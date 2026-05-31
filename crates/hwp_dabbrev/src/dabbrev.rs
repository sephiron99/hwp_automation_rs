use std::{
    cell::RefCell,
    collections::{HashMap, HashSet},
    rc::Rc,
};

use hwp_addon::{debug::log, text_edit::HwpEditExt};
use hwp_core::{
    hwp_obj::HwpObject,
    ihwpobject::lib::{GetTextStatus, ScanEpos, ScanRange, ScanSpos, mask},
};
use windows::Win32::UI::WindowsAndMessaging::GetForegroundWindow;

use crate::{AllTextCache, DabbrevPlugin, extract_all_words, ui_popup};

fn is_word_char(c: char) -> bool {
    c.is_alphanumeric() || c == '_' || c == '.' || c == '-' || c == ':'
}

fn is_match(word: &str, prefix: &str) -> bool {
    word.len() > prefix.len() && word.starts_with(prefix)
}

/// 현재 활성 문서의 캐시 키.
///
/// `HwpObject.Path`를 그대로 사용한다. 저장되지 않은 문서는 빈 문자열이
/// 반환되며, 모든 unsaved 문서는 같은 키 `""`로 통합된 캐시를 공유한다.
fn doc_key(hwp: &HwpObject) -> String {
    hwp.path().unwrap_or_default()
}

// ── Word Cache ──

#[derive(Default, Clone, Copy)]
struct WordStat {
    appearance: u32,
    chosen: u32,
}

impl WordStat {
    /// 채택빈도는 등장빈도의 10배 가중치.
    fn score(&self) -> u32 {
        self.appearance + 10 * self.chosen
    }
}

pub(crate) struct WordCache {
    entries: HashMap<String, WordStat>,
}

impl WordCache {
    fn new() -> Self {
        Self {
            entries: HashMap::new(),
        }
    }

    fn populate_from(&mut self, words: &[String]) {
        for w in words {
            self.entries.entry(w.clone()).or_default().appearance += 1;
        }
    }

    fn candidates(&self, prefix: &str) -> Vec<String> {
        let mut items: Vec<(&String, &WordStat)> = self
            .entries
            .iter()
            .filter(|(w, _)| is_match(w, prefix))
            .collect();
        items.sort_by(|a, b| b.1.score().cmp(&a.1.score()));
        items.into_iter().map(|(w, _)| w.clone()).collect()
    }

    fn record(&mut self, word: &str) {
        self.entries.entry(word.to_string()).or_default().chosen += 1;
    }
}

// ── State ──

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ExpansionPhase {
    /// next-word 모드 전용: all_text_cache에서 context_word 다음 단어를
    /// 한 번에 모두 수집하는 단계.
    AllTextCache,
    /// 캐시 소진 — 다음 fetch에서 rebuild 호출.
    NeedsReextract,
    /// 모든 소스 소진.
    Done,
}

pub(crate) struct DabbrevState {
    /// 이 state가 만들어진 문서의 키. 호출 시 doc_key가 달라지면 state는 폐기.
    doc_key: String,
    /// next-word 모드에서는 빈 문자열.
    prefix: String,
    /// next-word 모드일 때의 직전 단어.
    context_word: Option<String>,
    candidates: Vec<String>,
    current_index: usize,
    seen: HashSet<String>,
    phase: ExpansionPhase,
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

    /// 커서 바로 앞의 `prefix` 글자들을 `replacement`로 교체합니다.
    fn replace_word(
        &self,
        hwp: &HwpObject,
        prefix: &str,
        replacement: &str,
    ) -> hwp_core::error::Result<()> {
        log(
            "dabbrev",
            &format!("replace_word: prefix={prefix:?} -> {replacement:?}"),
        );
        hwp.replace_word_before(prefix, replacement)
    }

    /// 지정 문서(key)의 캐시가 없으면 text_segments()로 부트스트랩한다.
    fn bootstrap_if_needed(&self, hwp: &HwpObject, key: &str) -> hwp_core::error::Result<()> {
        let exists = self
            .word_caches
            .borrow()
            .as_ref()
            .is_some_and(|m| m.contains_key(key));
        if exists {
            return Ok(());
        }
        let words = extract_all_words(hwp)?;
        log(
            "dabbrev",
            &format!("bootstrap[{key:?}]: {} words", words.len()),
        );
        let mut wc = WordCache::new();
        wc.populate_from(&words);

        self.word_caches
            .borrow_mut()
            .get_or_insert_with(HashMap::new)
            .insert(key.to_string(), wc);

        let mut atc_caches = self.all_text_caches.borrow_mut();
        let map = atc_caches.get_or_insert_with(HashMap::new);
        map.entry(key.to_string())
            .or_insert_with(AllTextCache::new)
            .words = words;
        Ok(())
    }

    /// phase에 따라 후보를 추가로 채운다. 새 후보가 하나라도 들어오면 true.
    fn fetch_more(&self, hwp: &HwpObject, s: &mut DabbrevState) -> hwp_core::error::Result<bool> {
        let prev_len = s.candidates.len();
        let key = s.doc_key.clone();
        loop {
            match s.phase {
                ExpansionPhase::AllTextCache => {
                    // next-word 전용: cache.words에서 context_word 바로 뒤에 나오는
                    // 단어를 등장 순서대로(중복 제거) 한 번에 모두 모은다.
                    {
                        let caches = self.all_text_caches.borrow();
                        if let Some(cache) = caches.as_ref().and_then(|m| m.get(&key))
                            && let Some(ref cw) = s.context_word
                        {
                            for pair in cache.words.windows(2) {
                                if &pair[0] == cw && s.seen.insert(pair[1].clone()) {
                                    s.candidates.push(pair[1].clone());
                                }
                            }
                        }
                    }
                    s.phase = ExpansionPhase::NeedsReextract;
                    if s.candidates.len() > prev_len {
                        return Ok(true);
                    }
                    // 캐시에 후보가 전혀 없으면 같은 호출에서 곧장 rebuild 단계로.
                }
                ExpansionPhase::NeedsReextract => {
                    let added = {
                        let mut caches = self.all_text_caches.borrow_mut();
                        let map = caches.get_or_insert_with(HashMap::new);
                        map.entry(key.clone())
                            .or_insert_with(AllTextCache::new)
                            .rebuild(hwp)?
                    };
                    log(
                        "dabbrev",
                        &format!("reextract[{key:?}]: {} added words", added.len()),
                    );
                    let caches = self.all_text_caches.borrow();
                    let cache = caches.as_ref().and_then(|m| m.get(&key));
                    if let Some(ref cw) = s.context_word {
                        // next-word 모드: 새로 들어온 단어가 cw 뒤에 나오는 경우만 후보.
                        if let Some(cache) = cache {
                            let added_set: HashSet<&str> =
                                added.iter().map(|s| s.as_str()).collect();
                            for pair in cache.words.windows(2) {
                                if &pair[0] == cw && added_set.contains(pair[1].as_str()) {
                                    let cand = pair[1].clone();
                                    if s.seen.insert(cand.clone()) {
                                        s.candidates.push(cand);
                                    }
                                }
                            }
                        }
                    } else {
                        // prefix 모드: rebuild 후 전체 cache.words에서
                        // prefix 매칭 + 아직 seen에 안 들어간 것 모두 추가 (spec).
                        if let Some(cache) = cache {
                            for w in &cache.words {
                                if is_match(w, &s.prefix) && s.seen.insert(w.clone()) {
                                    s.candidates.push(w.clone());
                                }
                            }
                        }
                    }
                    s.phase = ExpansionPhase::Done;
                    return Ok(s.candidates.len() > prev_len);
                }
                ExpansionPhase::Done => return Ok(s.candidates.len() > prev_len),
            }
        }
    }

    pub fn expand(&self, hwp: &HwpObject) -> hwp_core::error::Result<bool> {
        let key = doc_key(hwp);
        self.bootstrap_if_needed(hwp, &key)?;

        let (line, char_after) = self.read_caret_context(hwp)?;
        // HWP GetText는 trailing 제어 문자(paragraph marker 등)를 포함할 수 있다.
        let line = line.trim_end_matches(|c: char| c.is_control()).to_string();
        let at_word = line.chars().last().is_some_and(is_word_char);

        let prefix = if at_word {
            line.rsplit(|c: char| !is_word_char(c))
                .next()
                .filter(|w| !w.is_empty())
                .map(|w| w.to_string())
        } else {
            None
        };

        log(
            "dabbrev",
            &format!(
                "doc_key={key:?}, prefix={prefix:?}, at_word={at_word}, char_after={char_after:?}"
            ),
        );

        let mut state = self.state.borrow_mut();

        // 문서가 바뀌었으면 state 폐기 (다른 문서의 cycle을 이어갈 수 없음).
        if let Some(ref s) = *state
            && s.doc_key != key
        {
            log("dabbrev", "doc switched — discarding state");
            *state = None;
        }

        // 연속 호출: 현재 line의 끝 단어가 직전 확장 결과와 일치하면 다음 후보로 cycle.
        if let Some(ref mut s) = *state {
            if let Some(ref p) = prefix
                && s.candidates.get(s.current_index) == Some(p)
            {
                let next = s.current_index + 1;
                if next < s.candidates.len() {
                    s.current_index = next;
                } else if self.fetch_more(hwp, s)? {
                    s.current_index = next;
                } else {
                    s.current_index = 0;
                }
                let word = s.candidates[s.current_index].clone();
                let prev_word = p.clone();
                log(
                    "dabbrev",
                    &format!(
                        "cycle -> {word:?} ({}/{})",
                        s.current_index + 1,
                        s.candidates.len()
                    ),
                );
                drop(state);
                self.replace_word(hwp, &prev_word, &word)?;
                self.run_popup_session(hwp, &word, true)?;
                return Ok(true);
            }
            // cycle이 아니다 — 이전 확장은 확정된 것으로 보고 채택빈도 기록.
            if s.context_word.is_none()
                && let Some(accepted) = s.candidates.get(s.current_index).cloned()
            {
                self.word_caches
                    .borrow_mut()
                    .get_or_insert_with(HashMap::new)
                    .entry(key.clone())
                    .or_insert_with(WordCache::new)
                    .record(&accepted);
                log("dabbrev", &format!("recorded chosen: {accepted:?}"));
            }
        }

        if let Some(prefix) = prefix {
            // ── prefix 모드 ──
            let mut seen = HashSet::new();
            let mut candidates: Vec<String> = Vec::new();

            // 1. word_cache (score 내림차순)
            let cached: Vec<String> = self
                .word_caches
                .borrow()
                .as_ref()
                .and_then(|m| m.get(&key))
                .map(|c| c.candidates(&prefix))
                .unwrap_or_default();
            for w in cached {
                if seen.insert(w.clone()) {
                    candidates.push(w);
                }
            }

            let mut s = DabbrevState {
                doc_key: key.clone(),
                prefix: prefix.clone(),
                context_word: None,
                candidates,
                current_index: 0,
                seen,
                // prefix 모드: word_cache로 prefix 매칭은 모두 제시되므로,
                // 추가 요청 시 곧장 rebuild로 이행 (순차 스캔 단계 없음).
                phase: ExpansionPhase::NeedsReextract,
            };

            // 캐시조차 비어있으면 후속 phase까지 즉시 밀어본다.
            while s.candidates.is_empty() && self.fetch_more(hwp, &mut s)? {}
            if s.candidates.is_empty() {
                *state = None;
                log("dabbrev", "매칭 없음");
                return Ok(true);
            }

            let expansion = s.candidates[0].clone();
            log(
                "dabbrev",
                &format!(
                    "expand {prefix:?} -> {expansion:?} (1/{})",
                    s.candidates.len()
                ),
            );
            self.replace_word(hwp, &prefix, &expansion)?;
            *state = Some(s);
            drop(state);
            self.run_popup_session(hwp, &expansion, true)?;
            Ok(true)
        } else {
            // ── next-word 모드 ──
            // 캐럿 바로 뒤에 단어가 붙어 있으면(단어 중간) next-word 확장하지 않는다.
            if char_after.is_some_and(is_word_char) {
                *state = None;
                log("dabbrev", "next-word: 캐럿 뒤가 단어 — 억제");
                return Ok(false);
            }
            let trimmed = line.trim_end_matches(|c: char| !is_word_char(c));
            let prev_word = trimmed
                .rsplit(|c: char| !is_word_char(c))
                .next()
                .filter(|w| !w.is_empty())
                .map(|w| w.to_string());

            let prev_word = match prev_word {
                Some(w) => w,
                None => {
                    *state = None;
                    return Ok(false);
                }
            };

            log("dabbrev", &format!("next-word mode: context={prev_word:?}"));

            let mut s = DabbrevState {
                doc_key: key.clone(),
                prefix: String::new(),
                context_word: Some(prev_word.clone()),
                candidates: Vec::new(),
                current_index: 0,
                seen: HashSet::new(),
                phase: ExpansionPhase::AllTextCache,
            };

            while s.candidates.is_empty() && self.fetch_more(hwp, &mut s)? {}
            if s.candidates.is_empty() {
                *state = None;
                log("dabbrev", "next-word: 매칭 없음");
                return Ok(true);
            }

            let expansion = s.candidates[0].clone();
            log(
                "dabbrev",
                &format!(
                    "next-word {prev_word:?} -> {expansion:?} (1/{})",
                    s.candidates.len()
                ),
            );
            *state = Some(s);
            drop(state);

            hwp.insert_text(&expansion)?;
            self.run_popup_session(hwp, &expansion, false)?;
            Ok(true)
        }
    }

    /// popup 세션 실행. self.state는 Some이어야 한다 (호출자가 보장).
    ///
    /// `initial_inserted` — popup 진입 직전 호출자가 문서에 넣은 첫 candidate.
    /// `mode_is_prefix` — prefix 모드인지(true), next-word 모드인지(false).
    /// 종료 후 state는 항상 None으로 정리된다.
    fn run_popup_session(
        &self,
        hwp: &HwpObject,
        initial_inserted: &str,
        mode_is_prefix: bool,
    ) -> hwp_core::error::Result<()> {
        let state = match self.state.borrow_mut().take() {
            Some(s) => s,
            None => return Ok(()),
        };
        let start_index = state.current_index;
        let initial = state.candidates.clone();
        let original_prefix = state.prefix.clone();
        let session_doc_key = state.doc_key.clone();

        // SAFETY: ui_popup::show()는 동기 message pump이며 plugin/hwp 참조는
        // 이 함수의 호출 동안 유효하다. 콜백은 popup 내부에서만 invocation됨.
        let plugin_ptr: *const Self = self;
        let hwp_ptr: *const HwpObject = hwp;
        let state_cell: Rc<RefCell<DabbrevState>> = Rc::new(RefCell::new(state));
        let last_inserted: Rc<RefCell<String>> =
            Rc::new(RefCell::new(initial_inserted.to_string()));

        let fetch_state = state_cell.clone();
        let fetch_cb = move || -> Vec<String> {
            let mut s = fetch_state.borrow_mut();
            let before = s.candidates.len();
            let plugin = unsafe { &*plugin_ptr };
            let hwp = unsafe { &*hwp_ptr };
            let _ = plugin.fetch_more(hwp, &mut s);
            s.candidates[before..].to_vec()
        };

        let replace_state = state_cell.clone();
        let replace_last = last_inserted.clone();
        let replace_cb = move |sel: usize| {
            let new_word = replace_state.borrow().candidates.get(sel).cloned();
            if let Some(new_word) = new_word {
                let hwp = unsafe { &*hwp_ptr };
                let prev = replace_last.borrow().clone();
                let _ = hwp.replace_word_before(&prev, &new_word);
                *replace_last.borrow_mut() = new_word;
                replace_state.borrow_mut().current_index = sel;
            }
        };

        let forward_target = unsafe { GetForegroundWindow() };
        let outcome = ui_popup::show(&initial, start_index, fetch_cb, replace_cb, forward_target);

        let last = last_inserted.borrow().clone();
        match outcome {
            ui_popup::Outcome::Committed => {
                log("dabbrev", &format!("popup committed: {last:?}"));
                self.word_caches
                    .borrow_mut()
                    .get_or_insert_with(HashMap::new)
                    .entry(session_doc_key)
                    .or_insert_with(WordCache::new)
                    .record(&last);
                *self.state.borrow_mut() = None;
            }
            ui_popup::Outcome::Cancelled => {
                log(
                    "dabbrev",
                    &format!("popup cancelled: restoring from {last:?}"),
                );
                let restore_to = if mode_is_prefix {
                    original_prefix.as_str()
                } else {
                    ""
                };
                let _ = hwp.replace_word_before(&last, restore_to);
                *self.state.borrow_mut() = None;
            }
        }
        Ok(())
    }
}
