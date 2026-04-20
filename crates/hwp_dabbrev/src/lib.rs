// 단어 매칭 순서 — 점진 스캔
//
// 1. prefix가 없는 경우, 직전 단어 다음에 나오는 단어를 찾아서 3, 4, 5, 6, 7번을 기준으로 순차대로 제시
// 2. word cache에서 기존 prefix에 매칭되었던 이력이 많았던 단어들을 빈도순 제시
// 3. 커서가 있는 문단과 그 이전 1개 문단에서 후보자 역순 제시
// 4. 커서가 있는 문단과 그 다음 1개 문단에서 후보자 순차 제시
// 5. 2번에서 탐색한 문단의 이전 1개 문단에서 후보자 역순 제시
// 6. 3번에서 탐색한 문단의 다음 1개 문단에서 후보자 순차 제시
// 7. 4번, 5번 반복
//
// 사용자가 확정하면 스캔 중단, word cache에 빈도 기록.
// 후보 소진 시에만 다음 문단을 추가 스캔.

use std::collections::{HashMap, HashSet};
use std::sync::{LazyLock, Mutex};

use hwp_addon::debug::log;
use hwp_addon::export_hwp_addon;
use hwp_addon::hwp_user_action::{ActionMeta, HwpUserAction, ToolbarConfig, ToolbarTarget};
use hwp_addon::shortcut::{Modifiers, ShortcutKey};
use hwp_addon::text_edit::HwpEditExt;
use hwp_core::hwp_obj::HwpObject;
use hwp_core::ihwpobject::lib::{mask, GetTextStatus, ScanEpos, ScanRange, ScanSpos};
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

/// 텍스트에서 단어를 추출합니다.
fn extract_words(text: &str) -> impl Iterator<Item = &str> {
    text.split(|c: char| !is_word_char(c))
        .filter(|w| !w.is_empty())
}

fn is_match(word: &str, prefix: &str) -> bool {
    word.len() > prefix.len() && word.starts_with(prefix)
}

/// 텍스트에서 `context_word` 바로 다음에 오는 단어들을 추출합니다.
fn extract_next_words<'a>(text: &'a str, context_word: &str) -> Vec<&'a str> {
    let words: Vec<&str> = extract_words(text).collect();
    words
        .windows(2)
        .filter(|pair| pair[0] == context_word)
        .map(|pair| pair[1])
        .collect()
}

/// 문단 단위로 점진적 스캔합니다.
///
/// `skip`개 문단을 건너뛰고 `collect`개 문단의 매칭 단어를 수집합니다.
/// `reverse_words`가 true이면 각 문단 내 단어를 역순으로 반환합니다 (후방 스캔용).
///
/// 반환: `(수집된 단어, 스캔 소진 여부)`
fn scan_paras(
    hwp: &HwpObject,
    prefix: &str,
    range: ScanRange,
    skip: usize,
    collect: usize,
    reverse_words: bool,
) -> hwp_core::error::Result<(Vec<String>, bool)> {
    hwp.init_scan(mask::NORMAL, range, 0, 0, 0, 0)?;
    let target_start = skip;
    let target_end = skip + collect;
    let mut para_index = 0;
    let mut para_words: Vec<String> = Vec::new();
    let mut all_words: Vec<String> = Vec::new();
    let mut exhausted = false;

    loop {
        let (status, text) = hwp.get_text()?;
        log(
            "scan_paras",
            &format!("para={para_index}, status={status:?}, text={text:?}"),
        );
        // 텍스트 수집을 문단 전환보다 먼저 수행
        match status {
            GetTextStatus::Normal | GetTextStatus::NextParagraph => {
                if para_index >= target_start {
                    para_words.extend(
                        extract_words(&text)
                            .filter(|w| is_match(w, prefix))
                            .map(|w| w.to_string()),
                    );
                }
            }
            GetTextStatus::EndOfList | GetTextStatus::None => {
                exhausted = true;
                break;
            }
            _ => {}
        }
        if status == GetTextStatus::NextParagraph {
            if para_index >= target_start && !para_words.is_empty() {
                if reverse_words {
                    para_words.reverse();
                }
                all_words.append(&mut para_words);
            }
            para_words.clear();
            para_index += 1;
            if para_index >= target_end {
                break;
            }
        }
    }
    if para_index >= target_start && !para_words.is_empty() {
        if reverse_words {
            para_words.reverse();
        }
        all_words.append(&mut para_words);
    }
    hwp.release_scan()?;
    log(
        "scan_paras",
        &format!("result: {} words, exhausted={exhausted}", all_words.len()),
    );
    Ok((all_words, exhausted))
}

/// `context_word` 다음에 오는 단어를 문단 단위로 점진적 스캔합니다.
///
/// `scan_paras`와 동일한 구조이나, prefix 매칭 대신
/// `context_word` 뒤에 오는 단어를 수집합니다.
fn scan_next_word_paras(
    hwp: &HwpObject,
    context_word: &str,
    range: ScanRange,
    skip: usize,
    collect: usize,
    reverse_words: bool,
) -> hwp_core::error::Result<(Vec<String>, bool)> {
    hwp.init_scan(mask::NORMAL, range, 0, 0, 0, 0)?;
    let target_start = skip;
    let target_end = skip + collect;
    let mut para_index = 0;
    let mut para_text = String::new();
    let mut all_words: Vec<String> = Vec::new();
    let mut exhausted = false;

    loop {
        let (status, text) = hwp.get_text()?;
        match status {
            GetTextStatus::Normal => {
                if para_index >= target_start {
                    para_text.push_str(&text);
                }
            }
            GetTextStatus::NextParagraph => {
                if para_index >= target_start {
                    para_text.push_str(&text);
                    let mut para_words: Vec<String> = extract_next_words(&para_text, context_word)
                        .into_iter()
                        .map(|w| w.to_string())
                        .collect();
                    if !para_words.is_empty() {
                        log(
                            "scan_next_word",
                            &format!("para={para_index}, matches={para_words:?}"),
                        );
                    }
                    if reverse_words {
                        para_words.reverse();
                    }
                    all_words.append(&mut para_words);
                }
                para_text.clear();
                para_index += 1;
                if para_index >= target_end {
                    break;
                }
            }
            GetTextStatus::EndOfList | GetTextStatus::None => {
                exhausted = true;
                break;
            }
            _ => {}
        }
    }
    if para_index >= target_start && !para_text.is_empty() {
        let mut para_words: Vec<String> = extract_next_words(&para_text, context_word)
            .into_iter()
            .map(|w| w.to_string())
            .collect();
        if !para_words.is_empty() {
            log(
                "scan_next_word",
                &format!("para={para_index}(tail), matches={para_words:?}"),
            );
        }
        if reverse_words {
            para_words.reverse();
        }
        all_words.append(&mut para_words);
    }
    hwp.release_scan()?;
    Ok((all_words, exhausted))
}

// ── Word Cache ──

/// prefix별 단어 선택 빈도를 추적하는 캐시
struct WordCache {
    /// prefix → (word → 선택 횟수)
    entries: HashMap<String, HashMap<String, u32>>,
}

impl WordCache {
    fn new() -> Self {
        Self {
            entries: HashMap::new(),
        }
    }

    /// 해당 prefix에 대해 캐시된 후보를 빈도 내림차순으로 반환합니다.
    fn candidates(&self, prefix: &str) -> Vec<String> {
        let Some(words) = self.entries.get(prefix) else {
            return Vec::new();
        };
        let mut items: Vec<_> = words.iter().collect();
        items.sort_by(|a, b| b.1.cmp(a.1));
        items.into_iter().map(|(w, _)| w.clone()).collect()
    }

    /// 사용자가 prefix에 대해 word를 확정했음을 기록합니다.
    fn record(&mut self, prefix: &str, word: &str) {
        *self
            .entries
            .entry(prefix.to_string())
            .or_default()
            .entry(word.to_string())
            .or_default() += 1;
    }
}

static WORD_CACHE: LazyLock<Mutex<WordCache>> = LazyLock::new(|| Mutex::new(WordCache::new()));

// ── State ──

struct DabbrevState {
    prefix: String,
    /// next-word 모드일 때의 문맥 단어 (직전 단어)
    context_word: Option<String>,
    candidates: Vec<String>,
    current_index: usize,
    seen: HashSet<String>,
    // 스캔 진행 상태
    backward_paras_done: usize,
    backward_exhausted: bool,
    forward_paras_done: usize,
    forward_exhausted: bool,
}

static STATE: Mutex<Option<DabbrevState>> = Mutex::new(None);

// ── Plugin ──

pub struct DabbrevPlugin;

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
    ///
    /// 프레임워크의 [`commit_composition`]이 액션 진입 전에 IME 조합을
    /// 커밋해놓으므로 prefix는 이미 문서에 실제로 존재한다. 실제 select+insert
    /// 패턴은 [`HwpEditExt::replace_word_before`]에 캡슐화되어 있어, 여기서는
    /// 로깅만 추가한다.
    ///
    /// [`commit_composition`]: hwp_addon::ime::commit_composition
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

    /// 후방/전방 각 1개 문단을 추가 스캔하여 새 후보를 찾습니다.
    ///
    /// 새 후보를 찾을 때까지 반복하며, 양방향 모두 소진되면 false를 반환합니다.
    /// `context_word`가 있으면 next-word 모드, 없으면 prefix 모드로 스캔합니다.
    /// 최대 `MAX_SCAN_ROUNDS`만큼만 스캔하여 HWP 응답없음을 방지합니다.
    fn scan_more(hwp: &HwpObject, s: &mut DabbrevState) -> hwp_core::error::Result<bool> {
        const MAX_SCAN_ROUNDS: usize = 10;
        for _ in 0..MAX_SCAN_ROUNDS {
            if s.backward_exhausted && s.forward_exhausted {
                return Ok(false);
            }
            let prev_len = s.candidates.len();

            // Step 4/6: 후방 1개 문단
            if !s.backward_exhausted {
                let range = ScanRange::new(ScanSpos::Document, ScanEpos::Current).backward();
                let (words, exhausted) = if let Some(ref cw) = s.context_word {
                    scan_next_word_paras(hwp, cw, range, s.backward_paras_done, 1, true)?
                } else {
                    scan_paras(hwp, &s.prefix, range, s.backward_paras_done, 1, true)?
                };
                s.backward_paras_done += 1;
                s.backward_exhausted = exhausted;
                for w in words {
                    if s.seen.insert(w.clone()) {
                        s.candidates.push(w);
                    }
                }
            }

            // Step 5/6: 전방 1개 문단
            if !s.forward_exhausted {
                let range = ScanRange::new(ScanSpos::Current, ScanEpos::Document);
                let (words, exhausted) = if let Some(ref cw) = s.context_word {
                    scan_next_word_paras(hwp, cw, range, s.forward_paras_done, 1, false)?
                } else {
                    scan_paras(hwp, &s.prefix, range, s.forward_paras_done, 1, false)?
                };
                s.forward_paras_done += 1;
                s.forward_exhausted = exhausted;
                for w in words {
                    if s.seen.insert(w.clone()) {
                        s.candidates.push(w);
                    }
                }
            }

            if s.candidates.len() > prev_len {
                return Ok(true);
            }
        }
        Ok(false)
    }

    fn expand(&self, hwp: &HwpObject) -> hwp_core::error::Result<bool> {
        let line = self.get_line_to_cursor(hwp)?;
        // HWP GetText는 trailing 제어 문자(paragraph marker 등)를 포함할 수 있으므로 제거
        let line = line.trim_end_matches(|c: char| c.is_control()).to_string();
        let at_word = line.chars().last().is_some_and(is_word_char);

        // prefix: 커서 바로 앞이 단어 문자이면 입력 중인 prefix
        let prefix = if at_word {
            line.rsplit(|c: char| !is_word_char(c))
                .next()
                .filter(|w| !w.is_empty())
                .map(|w| w.to_string())
        } else {
            None
        };

        log("dabbrev", &format!("prefix={prefix:?}, at_word={at_word}"));

        let mut state = STATE.lock().unwrap();

        // 연속 호출: 현재 단어가 마지막 확장 결과와 일치하면 다음 후보로
        if let Some(ref mut s) = *state {
            if let Some(ref p) = prefix {
                if s.candidates.get(s.current_index).map_or(false, |c| c == p) {
                    let next = s.current_index + 1;
                    if next < s.candidates.len() {
                        s.current_index = next;
                    } else if Self::scan_more(hwp, s)? {
                        s.current_index = next;
                    } else {
                        s.current_index = 0;
                    }
                    let word = s.candidates[s.current_index].clone();
                    // cycle 시에는 직전에 삽입한 후보(=현재 prefix p)를 교체합니다.
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
            // 이전 확장 확정 → 캐시에 기록 (prefix 모드일 때만)
            if s.context_word.is_none() {
                let accepted = s.candidates[s.current_index].clone();
                let prev_prefix = s.prefix.clone();
                WORD_CACHE.lock().unwrap().record(&prev_prefix, &accepted);
                log(
                    "dabbrev",
                    &format!("cache: {prev_prefix:?} -> {accepted:?}"),
                );
            }
        }

        if let Some(prefix) = prefix {
            // ── prefix 기반 확장 (기존 로직) ──

            let mut seen = HashSet::new();
            let mut candidates: Vec<String> = Vec::new();

            // 1. 캐시 (빈도순)
            for w in WORD_CACHE.lock().unwrap().candidates(&prefix) {
                if seen.insert(w.clone()) {
                    candidates.push(w);
                }
            }

            // 2. 후방 2개 문단 (커서 문단 + 이전 1개)
            let (bw, b_exhausted) = scan_paras(
                hwp,
                &prefix,
                ScanRange::new(ScanSpos::Document, ScanEpos::Current).backward(),
                0,
                2,
                true,
            )?;
            for w in bw {
                if seen.insert(w.clone()) {
                    candidates.push(w);
                }
            }

            // 3. 전방 2개 문단 (커서 문단 + 다음 1개)
            let (fw, f_exhausted) = scan_paras(
                hwp,
                &prefix,
                ScanRange::new(ScanSpos::Current, ScanEpos::Document),
                0,
                2,
                false,
            )?;
            for w in fw {
                if seen.insert(w.clone()) {
                    candidates.push(w);
                }
            }

            if candidates.is_empty() {
                // 초기 2개 문단에서 매칭 없음 — 소진되지 않았으면 더 깊이 스캔
                let mut s = DabbrevState {
                    prefix: prefix.clone(),
                    context_word: None,
                    candidates,
                    current_index: 0,
                    seen,
                    backward_paras_done: 2,
                    backward_exhausted: b_exhausted,
                    forward_paras_done: 2,
                    forward_exhausted: f_exhausted,
                };
                if !Self::scan_more(hwp, &mut s)? {
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
                return Ok(true);
            }

            let expansion = candidates[0].clone();
            log(
                "dabbrev",
                &format!(
                    "expand {prefix:?} -> {expansion:?} (1/{})",
                    candidates.len()
                ),
            );
            // replace_word가 prefix의 참조를 필요로 하므로 state 이동 전에 호출.
            self.replace_word(hwp, &prefix, &expansion)?;
            *state = Some(DabbrevState {
                prefix,
                context_word: None,
                candidates,
                current_index: 0,
                seen,
                backward_paras_done: 2,
                backward_exhausted: b_exhausted,
                forward_paras_done: 2,
                forward_exhausted: f_exhausted,
            });
            drop(state);
            Ok(true)
        } else {
            // ── next-word 모드: 직전 단어 다음에 오는 단어로 자동완성 ──

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

            let mut seen = HashSet::new();
            let mut candidates: Vec<String> = Vec::new();

            // 후방 2개 문단
            let (bw, b_exhausted) = scan_next_word_paras(
                hwp,
                &prev_word,
                ScanRange::new(ScanSpos::Document, ScanEpos::Current).backward(),
                0,
                2,
                true,
            )?;
            for w in bw {
                if seen.insert(w.clone()) {
                    candidates.push(w);
                }
            }

            // 전방 2개 문단
            let (fw, f_exhausted) = scan_next_word_paras(
                hwp,
                &prev_word,
                ScanRange::new(ScanSpos::Current, ScanEpos::Document),
                0,
                2,
                false,
            )?;
            for w in fw {
                if seen.insert(w.clone()) {
                    candidates.push(w);
                }
            }

            if candidates.is_empty() {
                *state = None;
                log("dabbrev", "next-word: 매칭 없음");
                return Ok(true);
            }

            let expansion = candidates[0].clone();
            log(
                "dabbrev",
                &format!(
                    "next-word {prev_word:?} -> {expansion:?} (1/{})",
                    candidates.len()
                ),
            );
            *state = Some(DabbrevState {
                prefix: String::new(),
                context_word: Some(prev_word),
                candidates,
                current_index: 0,
                seen,
                backward_paras_done: 2,
                backward_exhausted: b_exhausted,
                forward_paras_done: 2,
                forward_exhausted: f_exhausted,
            });
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

export_hwp_addon!(DabbrevPlugin, DabbrevPlugin);
