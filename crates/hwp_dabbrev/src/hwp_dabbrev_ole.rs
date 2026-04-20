//! hwp_dabbrev OLE 클라이언트 버전
//!
//! 실행 중인 HWP에 OLE로 연결하여 dabbrev 자동완성을 제공합니다.
//! `Ctrl+/` 전역 단축키로 트리거됩니다.
//!
//! # 실행 방법
//! 1. 한글을 먼저 실행한다.
//! 2. `cargo run -p hwp_dabbrev --bin hwp_dabbrev` 실행.
//! 3. 한글 문서에서 단어를 입력하다가 `Ctrl+/`를 누르면 dabbrev가 동작한다.
//! 4. 프로그램은 `Ctrl+C`로 종료한다.

use std::collections::{HashMap, HashSet};
use std::sync::{LazyLock, Mutex};

use hwp_core::hwp_obj::HwpObject;
use hwp_core::ihwpobject::lib::{GetTextStatus, ScanEpos, ScanRange, ScanSpos, mask};
use windows::Win32::System::Com::{CreateBindCtx, GetRunningObjectTable};
use windows::Win32::System::Ole::OleInitialize;
use windows::Win32::UI::Input::KeyboardAndMouse::{MOD_CONTROL, RegisterHotKey, VK_OEM_2};
use windows::Win32::UI::WindowsAndMessaging::{GetMessageW, MSG, WM_HOTKEY};

// ── 단어 판별 ──

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

/// 텍스트에서 `context_word` 바로 다음에 오는 단어들을 추출합니다.
fn extract_next_words<'a>(text: &'a str, context_word: &str) -> Vec<&'a str> {
    let words: Vec<&str> = extract_words(text).collect();
    words
        .windows(2)
        .filter(|pair| pair[0] == context_word)
        .map(|pair| pair[1])
        .collect()
}

// ── 스캔 ──

/// 문단 단위로 점진적 스캔합니다.
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
    Ok((all_words, exhausted))
}

/// `context_word` 다음에 오는 단어를 문단 단위로 점진적 스캔합니다.
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
    context_word: Option<String>,
    candidates: Vec<String>,
    current_index: usize,
    seen: HashSet<String>,
    backward_paras_done: usize,
    backward_exhausted: bool,
    forward_paras_done: usize,
    forward_exhausted: bool,
}

static STATE: Mutex<Option<DabbrevState>> = Mutex::new(None);

// ── 텍스트 편집 헬퍼 ──

/// 커서 위치까지의 문단 텍스트를 반환합니다.
fn get_line_to_cursor(hwp: &HwpObject) -> hwp_core::error::Result<String> {
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

/// 커서 바로 앞 `prefix` 길이만큼의 텍스트를 `replacement`로 교체합니다.
fn replace_word(hwp: &HwpObject, prefix: &str, replacement: &str) -> hwp_core::error::Result<()> {
    eprintln!("[dabbrev] replace: {prefix:?} -> {replacement:?}");
    let n = prefix.chars().count() as i32;
    if n > 0 {
        let (_, para, pos) = hwp.get_pos()?;
        let start = (pos - n).max(0);
        hwp.select_text(para, start, para, pos)?;
    }
    hwp.insert_text(replacement)
}

// ── Expand ──

/// 후방/전방 각 1개 문단을 추가 스캔하여 새 후보를 찾습니다.
fn scan_more(hwp: &HwpObject, s: &mut DabbrevState) -> hwp_core::error::Result<bool> {
    const MAX_SCAN_ROUNDS: usize = 10;
    for _ in 0..MAX_SCAN_ROUNDS {
        if s.backward_exhausted && s.forward_exhausted {
            return Ok(false);
        }
        let prev_len = s.candidates.len();

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

/// dabbrev 자동완성을 실행합니다.
fn expand(hwp: &HwpObject) -> hwp_core::error::Result<bool> {
    let line = get_line_to_cursor(hwp)?;
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

    eprintln!("[dabbrev] prefix={prefix:?}, at_word={at_word}");

    let mut state = STATE.lock().unwrap();

    // 연속 호출: 현재 단어가 마지막 확장 결과와 일치하면 다음 후보로
    if let Some(ref mut s) = *state {
        if let Some(ref p) = prefix {
            if s.candidates.get(s.current_index).map_or(false, |c| c == p) {
                let next = s.current_index + 1;
                if next < s.candidates.len() {
                    s.current_index = next;
                } else if scan_more(hwp, s)? {
                    s.current_index = next;
                } else {
                    s.current_index = 0;
                }
                let word = s.candidates[s.current_index].clone();
                let prev_word = p.clone();
                eprintln!(
                    "[dabbrev] cycle -> {word:?} ({}/{})",
                    s.current_index + 1,
                    s.candidates.len()
                );
                drop(state);
                replace_word(hwp, &prev_word, &word)?;
                return Ok(true);
            }
        }
        // 이전 확장 확정 → 캐시에 기록 (prefix 모드일 때만)
        if s.context_word.is_none() {
            let accepted = s.candidates[s.current_index].clone();
            let prev_prefix = s.prefix.clone();
            WORD_CACHE.lock().unwrap().record(&prev_prefix, &accepted);
            eprintln!("[dabbrev] cache: {prev_prefix:?} -> {accepted:?}");
        }
    }

    if let Some(prefix) = prefix {
        // ── prefix 기반 확장 ──

        let mut seen = HashSet::new();
        let mut candidates: Vec<String> = Vec::new();

        // 캐시 (빈도순)
        for w in WORD_CACHE.lock().unwrap().candidates(&prefix) {
            if seen.insert(w.clone()) {
                candidates.push(w);
            }
        }

        // 후방 2개 문단 (커서 문단 + 이전 1개)
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

        // 전방 2개 문단 (커서 문단 + 다음 1개)
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
            if !scan_more(hwp, &mut s)? {
                *state = None;
                eprintln!("[dabbrev] 매칭 없음");
                return Ok(true);
            }
            let expansion = s.candidates[0].clone();
            eprintln!(
                "[dabbrev] expand {prefix:?} -> {expansion:?} (1/{})",
                s.candidates.len()
            );
            replace_word(hwp, &prefix, &expansion)?;
            *state = Some(s);
            drop(state);
            return Ok(true);
        }

        let expansion = candidates[0].clone();
        eprintln!(
            "[dabbrev] expand {prefix:?} -> {expansion:?} (1/{})",
            candidates.len()
        );
        replace_word(hwp, &prefix, &expansion)?;
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

        eprintln!("[dabbrev] next-word mode: context={prev_word:?}");

        let mut seen = HashSet::new();
        let mut candidates: Vec<String> = Vec::new();

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
            eprintln!("[dabbrev] next-word: 매칭 없음");
            return Ok(true);
        }

        let expansion = candidates[0].clone();
        eprintln!(
            "[dabbrev] next-word {prev_word:?} -> {expansion:?} (1/{})",
            candidates.len()
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

// ── 진입점 ──

/// Running Object Table의 모든 moniker 이름을 출력합니다. (진단용)
fn dump_rot() -> windows::core::Result<()> {
    unsafe {
        OleInitialize(None)?;
        let rot = GetRunningObjectTable(0)?;
        let enum_moniker = rot.EnumRunning()?;
        let bind_ctx = CreateBindCtx(0)?;
        println!("── Running Object Table ──");
        let mut count = 0;
        loop {
            let mut moniker = [None];
            let mut fetched = 0u32;
            if enum_moniker.Next(&mut moniker, Some(&mut fetched)).is_err() || fetched == 0 {
                break;
            }
            if let Some(m) = &moniker[0] {
                match m.GetDisplayName(&bind_ctx, None) {
                    Ok(dn) => {
                        let name = dn.to_string().unwrap_or_default();
                        println!("  [{count}] {name}");
                    }
                    Err(e) => println!("  [{count}] <GetDisplayName 실패: {e}>"),
                }
            }
            count += 1;
        }
        println!("── 총 {count}개 ──");
    }
    Ok(())
}

fn main() -> hwp_com::Result<()> {
    let args: Vec<String> = std::env::args().collect();

    // 진단: --list-rot 인자로 실행 중인 COM 객체 목록 출력
    if args.iter().any(|a| a == "--list-rot") {
        dump_rot()?;
        return Ok(());
    }

    // HWP는 ROT에 자신을 자동 등록하지 않으므로 기본적으로 CoCreateInstance
    // 경로를 사용한다. HWP가 단일 인스턴스 COM 서버로 등록되어 있으면 기존
    // 창에 연결되고, 아니면 새 인스턴스가 뜬다.
    let use_attach = args.iter().any(|a| a == "--attach");

    let hwp = if use_attach {
        println!("HWP에 연결 중... (ROT 기반 attach)");
        match hwp_com::HwpClient::attach() {
            Ok(h) => h,
            Err(e) => {
                eprintln!("attach 실패: {e}");
                eprintln!("ROT 내용:");
                let _ = dump_rot();
                return Err(e);
            }
        }
    } else {
        println!("HWP 객체 생성 중... (CoCreateInstance)");
        let h = hwp_com::HwpClient::new()?;
        // 창을 확실히 표시 (이미 떠 있으면 no-op)
        if let Ok(windows) = h.windows() {
            if let Ok(active) = windows.active() {
                let _ = active.set_visible(true);
            }
        }
        h
    };
    println!("연결 완료. Ctrl+/를 눌러 dabbrev를 실행합니다. (Ctrl+C로 종료)");

    const HOTKEY_ID: i32 = 1;
    unsafe {
        RegisterHotKey(None, HOTKEY_ID, MOD_CONTROL, u32::from(VK_OEM_2.0))?;
    }

    let mut msg = MSG::default();
    loop {
        let result = unsafe { GetMessageW(&mut msg, None, 0, 0) };
        match result.0 {
            0 => break,  // WM_QUIT
            -1 => break, // 오류
            _ => {}
        }
        if msg.message == WM_HOTKEY && msg.wParam.0 as i32 == HOTKEY_ID {
            if let Err(e) = expand(&hwp) {
                eprintln!("[dabbrev] 오류: {e}");
            }
        }
    }

    Ok(())
}
