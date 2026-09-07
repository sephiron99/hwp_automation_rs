use std::collections::HashSet;

use hwp_addon::{debug::log, text_edit::HwpEditExt};
use hwp_core::{
    hwp_obj::HwpObject,
    ihwpobject::lib::{GetTextStatus, ScanEpos, ScanRange, ScanSpos, mask},
};
use winsafe::HWND;

use crate::{DabbrevPlugin, extract_all_words, is_word_char, strip_leading_nonword, ui_popup};

fn is_match(word: &str, prefix: &str) -> bool {
    word.len() > prefix.len() && word.starts_with(prefix)
}

/// 후보를 삽입한다. prefix가 있으면 교체하고, next-word 모드에서는 직전
/// 입력(space 등)과 undo 레코드가 병합되지 않도록 빈 범위를 선택한다.
fn insert_candidate(
    hwp: &HwpObject,
    prefix: Option<&str>,
    word: &str,
) -> hwp_core::error::Result<()> {
    if let Some(prefix) = prefix {
        hwp.replace_word_before(prefix, word)
    } else {
        let (_, para, pos) = hwp.get_pos()?;
        let _ = hwp.select_text(para, pos, para, pos)?;
        hwp.insert_text(word)
    }
}

impl DabbrevPlugin {
    /// 현재 문단을 읽어 `(캐럿 앞 텍스트, 캐럿 바로 뒤 글자)`를 반환합니다.
    ///
    /// `GetText`의 `ScanEpos::Current`가 캐럿에서 멈추지 않고 문단 끝까지 읽는
    /// 경우가 있어, prefix/모드 판정이 캐럿 뒤 텍스트로 오염된다. 그래서 전체
    /// 문단을 읽은 뒤 `GetPos`의 문단 내 문자 offset(`pos`)으로 직접 잘라낸다.
    fn read_caret_context(
        &self,
        hwp: &HwpObject,
    ) -> hwp_core::error::Result<(String, Option<char>)> {
        let (_, _, pos) = hwp.get_pos()?;
        let caret = pos.max(0) as usize;

        hwp.init_scan(
            mask::NORMAL,
            ScanRange::new(ScanSpos::Paragraph, ScanEpos::Paragraph),
            0,
            0,
            0,
            0,
        )?;
        let mut para = String::new();
        loop {
            let (status, text) = hwp.get_text()?;
            match status {
                GetTextStatus::Normal => para.push_str(&text),
                _ => break,
            }
        }
        hwp.release_scan()?;

        let before: String = para.chars().take(caret).collect();
        let after = para.chars().nth(caret);
        Ok((before, after))
    }

    pub fn expand(&self, hwp: &HwpObject) -> hwp_core::error::Result<bool> {
        let words = extract_all_words(hwp)?;
        let (line, char_after) = self.read_caret_context(hwp)?;
        // HWP GetText는 trailing 제어 문자(paragraph marker 등)를 포함할 수 있다.
        let line = line.trim_end_matches(|c: char| c.is_control());
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

        log(
            "dabbrev",
            &format!(
                "{} words, prefix={prefix:?}, at_word={at_word}, char_after={char_after:?}",
                words.len()
            ),
        );

        // 두 모드 모두 문서 등장 순서를 유지하며 중복 후보를 제거한다.
        let mut seen = HashSet::new();
        let candidates: Vec<String> = if let Some(prefix) = prefix.as_deref() {
            words
                .iter()
                .filter(|word| is_match(word, prefix))
                .filter(|word| seen.insert(word.as_str()))
                .cloned()
                .collect()
        } else {
            // 캐럿 바로 뒤에 단어가 붙어 있으면 next-word 확장하지 않는다.
            if char_after.is_some_and(is_word_char) {
                log("dabbrev", "next-word: 캐럿 뒤가 단어 — 억제");
                return Ok(false);
            }
            let trimmed = line.trim_end_matches(|c: char| !is_word_char(c));
            let Some(prev_word) = trimmed
                .rsplit(|c: char| !is_word_char(c))
                .next()
                .map(strip_leading_nonword)
                .filter(|w| !w.is_empty())
            else {
                return Ok(false);
            };

            log("dabbrev", &format!("next-word mode: context={prev_word:?}"));
            words
                .windows(2)
                .filter(|pair| pair[0] == prev_word)
                .map(|pair| &pair[1])
                .filter(|word| seen.insert(word.as_str()))
                .cloned()
                .collect()
        };

        if candidates.is_empty() {
            log("dabbrev", "매칭 없음");
            return Ok(true);
        }

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
    ) -> hwp_core::error::Result<()> {
        let Some(first) = candidates.first() else {
            return Ok(());
        };
        log(
            "dabbrev",
            &format!("expand {prefix:?} -> {first:?} (1/{})", candidates.len()),
        );
        insert_candidate(hwp, prefix.as_deref(), first)?;

        // winsafe 이벤트는 'static 클로저이므로 HwpObject와 prefix를 소유한다.
        // 후보를 바꿀 때마다 직전 삽입을 되돌려 undo 항목을 하나로 유지한다.
        let replace_hwp = hwp.clone();
        let replace_cb = move |word: &str| {
            let _ = replace_hwp.undo();
            let _ = insert_candidate(&replace_hwp, prefix.as_deref(), word);
        };

        let forward_target = HWND::GetForegroundWindow();
        match ui_popup::show(candidates, replace_cb, forward_target) {
            ui_popup::Outcome::Committed => log("dabbrev", "popup committed"),
            ui_popup::Outcome::Cancelled => {
                log("dabbrev", "popup cancelled");
                // prefix 모드 → 원래 prefix, next-word → 공백 뒤 원상태.
                let _ = hwp.undo();
            }
        }
        Ok(())
    }
}
