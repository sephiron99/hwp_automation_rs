use std::collections::HashSet;

use hwp_addon::{
    HwpExt, HwpObject,
    method::hwp_object::{ScanRange, mask},
    text_edit::HwpEditExt,
    text_extract::{HwpTextExt, TextSegment},
    variant::IntoVariant,
};
use winsafe::HWND;

use crate::{DabbrevPlugin, extract_all_words, is_word_char, strip_leading_nonword, ui_popup};

/// 문서에서 추출한 `word`가 입력 중인 `prefix`를 확장할 후보인지 판정한다.
///
/// 다음 두 조건을 모두 만족해야 한다.
/// - `word`가 `prefix`로 시작해야 한다. 중간에 포함된 문자열은 매칭하지 않는다.
/// - `prefix` 뒤에 추가 내용이 있어야 한다. 이미 입력한 단어와 같은 후보는
///   자동완성을 해도 내용이 바뀌지 않으므로 제외한다.
///
/// `str::len()`은 글자 수가 아닌 UTF-8 바이트 수를 반환한다. 여기서는
/// `starts_with`로 접두사 일치도 확인하므로, 더 긴 바이트 길이라는 조건만으로
/// 접두사 뒤에 내용이 더 있는지 판정할 수 있다. 한글도 같은 방식으로 처리한다.
///
/// 대소문자를 구분하며 유니코드 정규화나 한글 초성 검색은 하지 않는다.
/// 예를 들어 `("자동완성", "자동")`은 참이고, `("자동", "자동")`과
/// `("반자동", "자동")`, `("Rust", "rust")`는 거짓이다.
///
/// 빈 `prefix`에는 비어 있지 않은 모든 `word`가 매칭된다. 현재 호출부는
/// 빈 접두사를 미리 걸러내므로, 이 함수에는 실제 입력된 접두사만 전달한다.
fn is_match(word: &str, prefix: &str) -> bool {
    word.len() > prefix.len() && word.starts_with(prefix)
}

/// 후보를 삽입한다. prefix가 있으면 교체하고, next-word 모드에서는 직전
/// 입력(space 등)과 undo 레코드가 병합되지 않도록 빈 범위를 선택한다.
fn insert_candidate(hwp: &HwpObject, prefix: Option<&str>, word: &str) -> hwp_addon::Result<()> {
    if let Some(prefix) = prefix {
        hwp.replace_word_before(prefix, word)
    } else {
        let (_, para, pos) = hwp.get_pos()?;
        hwp.select_text(para, pos, para, pos)?;
        hwp.insert_text(word)
    }
}

impl DabbrevPlugin {
    /// 현재 캐럿 주변을 읽어 `(캐럿 앞 텍스트, 캐럿 바로 뒤 글자)`를 반환합니다.
    ///
    /// `MoveSelWordBegin`으로 앞쪽을 읽되, 공백·구두점만 선택된 경우에는
    /// 문단 시작까지 범위를 넓혀 next-word 모드에 필요한 직전 단어도 읽는다.
    /// 뒤쪽은 원래 위치에서 `MoveSelWordEnd`로 선택하여 첫 글자를 얻는다.
    /// 각 방향의 기준 위치와 최종 캐럿 위치는 저장한 좌표에 `SetPos`하여 복원한다.
    fn read_caret_context(&self, hwp: &HwpObject) -> hwp_addon::Result<(String, Option<char>)> {
        hwp.call::<()>("UnSelectCtrl")?;
        let action = hwp.h_action()?;
        action.run("Cancel")?;
        action.run("Cancel")?;
        let (list, para, pos) = hwp.get_pos()?;

        let restore_caret = || -> hwp_addon::Result<()> {
            hwp.call_with(
                "SetPos",
                vec![
                    list.into_variant()?,
                    para.into_variant()?,
                    pos.into_variant()?,
                ],
            )
        };
        let read_selection = || -> hwp_addon::Result<String> {
            // 문서 경계 등에서 선택 범위가 생기지 않았다면 스캔하지 않는다.
            if hwp.get_pos()? == (list, para, pos) {
                return Ok(String::new());
            }
            let mut text = String::new();
            for segment in
                hwp.text_segments_with(mask::NORMAL, ScanRange::within_selection(), 0, 0, 0, 0)?
            {
                match segment? {
                    TextSegment::Text(s) | TextSegment::EnterCtrl(s) => text.push_str(&s),
                    TextSegment::ParaBreak => text.push('\n'),
                    TextSegment::ExitCtrl => {}
                }
            }
            Ok(text)
        };

        let result: hwp_addon::Result<_> = (|| {
            action.run("MoveSelWordBegin")?;
            let mut before = read_selection()?;
            if pos > 0 && !before.chars().any(char::is_alphanumeric) {
                // "오늘 |"에서 공백만 읽으면 prev_word를 구할 수 없다.
                // 원래 캐럿을 기준으로 범위를 넓히고, 뒤의 공백은 그대로 보존한다.
                restore_caret()?;
                action.run("MoveSelParaBegin")?;
                before = read_selection()?;
                tracing::debug!("직전 단어 검색: 문단 시작까지 선택 범위 확장");
            }
            restore_caret()?;
            action.run("MoveSelWordEnd")?;
            let after = read_selection()?.chars().next();
            Ok((before, after))
        })();
        // 이동 또는 읽기가 실패해도 원래 위치로 복원을 시도한다.
        let restored = restore_caret();
        let context = result?;
        restored?;
        Ok(context)
    }

    pub fn expand(&self, hwp: &HwpObject) -> hwp_addon::Result<bool> {
        let words = extract_all_words(hwp)?;
        let (line, char_after) = self.read_caret_context(hwp)?;
        // 스캔 끝의 문단 표시는 제거하되, next-word 모드의 구분자인 탭은 보존한다.
        let line = line.trim_end_matches(['\r', '\n']);
        let at_word = line.chars().last().is_some_and(is_word_char);

        let prefix = if at_word {
            line.rsplit(|c: char| !is_word_char(c))
                .next()
                .map(strip_leading_nonword)
                .filter(|w| !w.is_empty())
                .map(str::to_string)
        } else {
            None
        };

        tracing::debug!(
            words = words.len(),
            ?prefix,
            at_word,
            ?char_after,
            "단어 수집"
        );

        // 두 모드 모두 문서 등장 순서를 유지하며 중복 후보를 제거한다.
        let mut seen = HashSet::new();
        let mut candidates: Vec<String> = if let Some(prefix) = prefix.as_deref() {
            words
                .iter()
                .filter(|word| is_match(word, prefix))
                .filter(|word| seen.insert(word.as_str()))
                .cloned()
                .collect()
        } else {
            // 캐럿 바로 뒤에 단어가 붙어 있으면 next-word 확장하지 않는다.
            if char_after.is_some_and(is_word_char) {
                tracing::debug!("next-word: 캐럿 뒤가 단어 — 억제");
                return Ok(false);
            }
            let trimmed = line.trim_end_matches(|c: char| !is_word_char(c));
            let Some(prev_word) = trimmed
                .rsplit(|c: char| !is_word_char(c))
                .next()
                .map(strip_leading_nonword)
                .filter(|w| !w.is_empty())
            else {
                tracing::debug!("next-word: 직전 단어 없음");
                return Ok(false);
            };

            tracing::debug!(context = ?prev_word, "next-word mode");
            words
                .windows(2)
                .filter(|pair| pair[0] == prev_word)
                .map(|pair| &pair[1])
                .filter(|word| seen.insert(word.as_str()))
                .cloned()
                .collect()
        };

        if candidates.is_empty() {
            tracing::debug!("매칭 없음");
            return Ok(true);
        }
        // 선택 이력이 있는 후보를 앞으로 옮긴다. 나머지는 문서 등장 순서를 유지한다.
        self.history.borrow().prioritize(&mut candidates);

        self.run_popup_session(hwp, prefix, candidates)?;
        Ok(true)
    }

    /// 완성된 후보 목록으로 popup을 실행한다. 목록과 선택 상태는 이 세션에서만
    /// 유지하며, 다음 액션에서는 문서에서 다시 구성한다.
    fn run_popup_session(
        &self,
        hwp: &HwpObject,
        prefix: Option<String>,
        candidates: Vec<String>,
    ) -> hwp_addon::Result<()> {
        let Some(first) = candidates.first() else {
            return Ok(());
        };
        tracing::debug!(?prefix, ?first, candidates = candidates.len(), "expand");
        insert_candidate(hwp, prefix.as_deref(), first)?;

        // winsafe 이벤트는 'static 클로저이므로 HwpObject와 prefix를 소유한다.
        // 후보를 바꿀 때마다 직전 삽입을 되돌려 undo 항목을 하나로 유지한다.
        let replace_hwp = hwp.clone();
        let replace_cb = move |word: &str| {
            let _ = replace_hwp.h_action().and_then(|a| a.undo());
            let _ = insert_candidate(&replace_hwp, prefix.as_deref(), word);
        };

        let forward_target = HWND::GetForegroundWindow();
        // show()가 후보 목록을 소유하므로 확정한 단어를 찾을 사본을 남긴다.
        match ui_popup::show(candidates.clone(), replace_cb, forward_target) {
            ui_popup::Outcome::Committed(index) => {
                tracing::debug!(chosen = ?candidates.get(index), "popup committed");
                if let Some(word) = candidates.get(index) {
                    self.history.borrow_mut().record(word);
                }
            }
            ui_popup::Outcome::Cancelled => {
                tracing::debug!("popup cancelled");
                // prefix 모드 → 원래 prefix, next-word → 공백 뒤 원상태.
                let _ = hwp.h_action().and_then(|a| a.undo());
            }
        }
        Ok(())
    }
}
