// word cache와 all_text_cache 두가지가 존재
// word cache는 채택빈도도 저장 (채택빈도 가중치 = 등장빈도의 100배)
// all_text_cache는 직전에 얻은 text_segments()의 결과를 단순히 캐싱
// 1. word cache가 처음 구성되는 경우 text_extract를 기반으로 캐시 구성
// 2. prefix가 없는 경우, all_text_cache에서 직전 단어 다음에 나오는 단어를 찾아서 순차 제시하고, all_text_cache의 후보를 다 소진했는데도 사용자가 계속 요구하는 경우, text_extract를 실행하여 all_text_cache를 업데이트하고 diff하여 달라진 부분만 매칭, 그리고 all_text_cache를 업데이트
// 3. prefix가 있는 경우, word_cache에서 채택빈도순으로 제시하고, 다 제시했는데도 없다면 all_text_cache에서 찾아서 순차 제시하고, all_text_cache의 후보를 다 소진했는데도 사용자가 계속 요구하는 경우, text_extract를 실행하여 all_text_cache와 diff하여 달라진 부분만 매칭, 그리고 all_text_cache를 업데이트
//
// 사용자가 확정하면, word cache에 빈도 기록.

use std::cell::RefCell;
use std::collections::{HashMap, HashSet};

use hwp_addon::debug::log;
use hwp_addon::export_hwp_addon;
use hwp_addon::hwp_user_action::{ActionMeta, HwpUserAction, ToolbarConfig, ToolbarTarget};
use hwp_addon::shortcut::{Modifiers, ShortcutKey};
use hwp_addon::text_edit::HwpEditExt;
use hwp_addon::text_extract::{HwpTextExt, TextSegment};
use hwp_core::hwp_obj::HwpObject;
use hwp_core::ihwpobject::lib::{GetTextStatus, ScanEpos, ScanRange, ScanSpos, mask};
use windows::Win32::UI::Input::KeyboardAndMouse::VK_OEM_2;

const TOOLBAR_DATA: &[u8] = include_bytes!("../toolbar.bmp");

const ACTION_DABBREV: &str = "DabbrevExpand";

static CONFIG: ToolbarConfig = ToolbarConfig {
    name: "dabbrev",
    serialize_path: "HwpDabbrev",
    bitmap_data: TOOLBAR_DATA,
    target: ToolbarTarget::Ribbon("자동완성"),
    ribbon_toolbox_index: -1,
};

fn is_word_char(c: char) -> bool {
    c.is_alphanumeric() || c == '_' || c == '.' || c == '-' || c == ':'
}

fn extract_words(text: &str) -> impl Iterator<Item = &str> {
    text.split(|c: char| !is_word_char(c))
        .filter(|w| !w.is_empty())
}

fn is_match(word: &str, prefix: &str) -> bool {
    word.len() > prefix.len() && word.starts_with(prefix)
}

/// `text_segments()`로 문서 전체를 훑어 단어 배열을 만든다. 컨트롤 진입/탈출
/// 경계는 무시하고 그 안의 텍스트도 본문과 같은 흐름으로 평탄화한다.
fn extract_all_words(hwp: &HwpObject) -> hwp_core::error::Result<Vec<String>> {
    let mut buf = String::new();
    for seg in hwp.text_segments()? {
        match seg? {
            TextSegment::Text(s) | TextSegment::EnterCtrl(s) => buf.push_str(&s),
            TextSegment::ParaBreak => buf.push(' '),
            TextSegment::ExitCtrl => {}
        }
    }
    Ok(extract_words(&buf).map(|w| w.to_string()).collect())
}

// ── Word Cache ──

#[derive(Default, Clone, Copy)]
struct WordStat {
    appearance: u32,
    chosen: u32,
}

impl WordStat {
    /// 채택빈도는 등장빈도의 100배 가중치.
    fn score(&self) -> u32 {
        self.appearance + 100 * self.chosen
    }
}

struct WordCache {
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

// ── All Text Cache ──

struct AllTextCache {
    words: Vec<String>,
}

impl AllTextCache {
    const fn new() -> Self {
        Self { words: Vec::new() }
    }

    /// 새 `text_segments()` 결과로 교체하고, 이전 `words`에 없던 단어를
    /// 등장 순서대로 (중복 제거하여) 반환한다. 첫 호출 시에는 전체 단어 set이
    /// 그대로 added로 간주된다.
    fn rebuild(&mut self, hwp: &HwpObject) -> hwp_core::error::Result<Vec<String>> {
        let new_words = extract_all_words(hwp)?;
        let old_set: HashSet<&str> = self.words.iter().map(|s| s.as_str()).collect();
        let mut added_set: HashSet<String> = HashSet::new();
        let mut added: Vec<String> = Vec::new();
        for w in &new_words {
            if !old_set.contains(w.as_str()) && added_set.insert(w.clone()) {
                added.push(w.clone());
            }
        }
        self.words = new_words;
        Ok(added)
    }
}

// ── State ──

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ExpansionPhase {
    /// all_text_cache 순차 스캔 진행 인덱스.
    AllTextCache(usize),
    /// 캐시 소진 — 다음 fetch에서 rebuild 호출.
    NeedsReextract,
    /// 모든 소스 소진.
    Done,
}

struct DabbrevState {
    /// next-word 모드에서는 빈 문자열.
    prefix: String,
    /// next-word 모드일 때의 직전 단어.
    context_word: Option<String>,
    candidates: Vec<String>,
    current_index: usize,
    seen: HashSet<String>,
    phase: ExpansionPhase,
}

// ── Plugin ──

pub struct DabbrevPlugin {
    state: RefCell<Option<DabbrevState>>,
    word_cache: RefCell<Option<WordCache>>,
    all_text_cache: RefCell<AllTextCache>,
}

impl DabbrevPlugin {
    /// 커서 위치까지의 문단 텍스트를 반환합니다.
    fn get_line_to_cursor(&self, hwp: &HwpObject) -> hwp_core::error::Result<String> {
        hwp.init_scan(
            mask::NORMAL,
            ScanRange::new(ScanSpos::Paragraph, ScanEpos::Current),
            0,
            0,
            0,
            0,
        )?;
        let mut line = String::new();
        loop {
            let (status, text) = hwp.get_text()?;
            match status {
                GetTextStatus::Normal => line.push_str(&text),
                _ => break,
            }
        }
        hwp.release_scan()?;
        Ok(line)
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

    /// 새로 부트스트랩이 필요한지 확인하고, 필요하면 text_segments()를 한 번 돌려
    /// `word_cache`와 `all_text_cache`를 동시에 초기화한다.
    fn bootstrap_if_needed(&self, hwp: &HwpObject) -> hwp_core::error::Result<()> {
        if self.word_cache.borrow().is_some() {
            return Ok(());
        }
        let words = extract_all_words(hwp)?;
        log("dabbrev", &format!("bootstrap: {} words", words.len()));
        let mut wc = WordCache::new();
        wc.populate_from(&words);
        *self.word_cache.borrow_mut() = Some(wc);
        self.all_text_cache.borrow_mut().words = words;
        Ok(())
    }

    /// phase에 따라 후보를 추가로 채운다. 새 후보가 하나라도 들어오면 true.
    fn fetch_more(&self, hwp: &HwpObject, s: &mut DabbrevState) -> hwp_core::error::Result<bool> {
        let prev_len = s.candidates.len();
        loop {
            match s.phase {
                ExpansionPhase::AllTextCache(mut idx) => {
                    let cache = self.all_text_cache.borrow();
                    let words = &cache.words;
                    if let Some(ref cw) = s.context_word {
                        // next-word: words[i] == cw 면 words[i+1]이 후보.
                        while idx + 1 < words.len() {
                            let i = idx;
                            idx += 1;
                            if &words[i] == cw {
                                let cand = words[i + 1].clone();
                                if s.seen.insert(cand.clone()) {
                                    s.candidates.push(cand);
                                    s.phase = ExpansionPhase::AllTextCache(idx);
                                    return Ok(true);
                                }
                            }
                        }
                    } else {
                        while idx < words.len() {
                            let i = idx;
                            idx += 1;
                            if is_match(&words[i], &s.prefix) {
                                let cand = words[i].clone();
                                if s.seen.insert(cand.clone()) {
                                    s.candidates.push(cand);
                                    s.phase = ExpansionPhase::AllTextCache(idx);
                                    return Ok(true);
                                }
                            }
                        }
                    }
                    drop(cache);
                    s.phase = ExpansionPhase::NeedsReextract;
                }
                ExpansionPhase::NeedsReextract => {
                    let added = self.all_text_cache.borrow_mut().rebuild(hwp)?;
                    log(
                        "dabbrev",
                        &format!("reextract: {} added words", added.len()),
                    );
                    if let Some(ref cw) = s.context_word {
                        // next-word 모드: added 자체만 보면 windows(2) 컨텍스트가
                        // 사라지므로, 새 words 전체에서 windows(2) 매칭하되
                        // "next 위치"가 added인 경우만 후보로 받는다.
                        let cache = self.all_text_cache.borrow();
                        let added_set: HashSet<&str> = added.iter().map(|s| s.as_str()).collect();
                        for pair in cache.words.windows(2) {
                            if &pair[0] == cw && added_set.contains(pair[1].as_str()) {
                                let cand = pair[1].clone();
                                if s.seen.insert(cand.clone()) {
                                    s.candidates.push(cand);
                                }
                            }
                        }
                    } else {
                        for w in &added {
                            if is_match(w, &s.prefix) && s.seen.insert(w.clone()) {
                                s.candidates.push(w.clone());
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

    fn expand(&self, hwp: &HwpObject) -> hwp_core::error::Result<bool> {
        self.bootstrap_if_needed(hwp)?;

        let line = self.get_line_to_cursor(hwp)?;
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

        log("dabbrev", &format!("prefix={prefix:?}, at_word={at_word}"));

        let mut state = self.state.borrow_mut();

        // 연속 호출: 현재 line의 끝 단어가 직전 확장 결과와 일치하면 다음 후보로 cycle.
        if let Some(ref mut s) = *state {
            if let Some(ref p) = prefix {
                if s.candidates.get(s.current_index).map_or(false, |c| c == p) {
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
                    return Ok(true);
                }
            }
            // cycle이 아니다 — 이전 확장은 확정된 것으로 보고 채택빈도 기록.
            if s.context_word.is_none() {
                if let Some(accepted) = s.candidates.get(s.current_index).cloned() {
                    self.word_cache
                        .borrow_mut()
                        .get_or_insert_with(WordCache::new)
                        .record(&accepted);
                    log("dabbrev", &format!("recorded chosen: {accepted:?}"));
                }
            }
        }

        if let Some(prefix) = prefix {
            // ── prefix 모드 ──
            let mut seen = HashSet::new();
            let mut candidates: Vec<String> = Vec::new();

            // 1. word_cache (score 내림차순)
            let cached: Vec<String> = self
                .word_cache
                .borrow()
                .as_ref()
                .map(|c| c.candidates(&prefix))
                .unwrap_or_default();
            for w in cached {
                if seen.insert(w.clone()) {
                    candidates.push(w);
                }
            }

            let mut s = DabbrevState {
                prefix: prefix.clone(),
                context_word: None,
                candidates,
                current_index: 0,
                seen,
                phase: ExpansionPhase::AllTextCache(0),
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
            Ok(true)
        } else {
            // ── next-word 모드 ──
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
                prefix: String::new(),
                context_word: Some(prev_word.clone()),
                candidates: Vec::new(),
                current_index: 0,
                seen: HashSet::new(),
                phase: ExpansionPhase::AllTextCache(0),
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
            Ok(true)
        }
    }
}

impl HwpUserAction for DabbrevPlugin {
    fn toolbar_config(&self) -> Option<&ToolbarConfig> {
        Some(&CONFIG)
    }

    fn actions(&self) -> &'static [ActionMeta] {
        static ACTIONS: [ActionMeta; 1] = [ActionMeta {
            name: ACTION_DABBREV,
            label: "dabbrev",
            image_index: 0,
            shortcut: Some(ShortcutKey {
                modifiers: Modifiers {
                    alt: false,
                    ctrl: true,
                    shift: false,
                },
                key: VK_OEM_2, // '/' 키
            }),
        }];
        &ACTIONS
    }

    fn on_load(&self, hwp: &HwpObject) -> hwp_core::error::Result<bool> {
        match self.setup_toolbar(hwp) {
            Ok(v) => Ok(v),
            Err(e) => {
                log("dabbrev:on_load", &format!("{e}"));
                Err(e)
            }
        }
    }

    fn do_action(&self, action_name: &str, hwp: &HwpObject) -> hwp_core::error::Result<bool> {
        match action_name {
            ACTION_DABBREV => match self.expand(hwp) {
                Ok(v) => Ok(v),
                Err(e) => {
                    log("dabbrev:expand", &format!("{e}"));
                    Err(e)
                }
            },
            _ => Ok(false),
        }
    }
}

export_hwp_addon!(
    DabbrevPlugin,
    DabbrevPlugin {
        state: RefCell::new(None),
        word_cache: RefCell::new(None),
        all_text_cache: RefCell::new(AllTextCache::new()),
    }
);
