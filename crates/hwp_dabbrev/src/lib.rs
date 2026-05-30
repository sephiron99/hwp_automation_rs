// word cache와 all_text_cache 두가지가 존재
// word cache는 채택빈도도 저장 (채택빈도 가중치 = 등장빈도의 100배)
// all_text_cache는 직전에 얻은 text_segments()의 결과를 단순히 캐싱
// 1. word cache가 처음 구성되는 경우 text_extract를 기반으로 캐시 구성
// 2. prefix가 없는 경우, all_text_cache에서 직전 단어 다음에 나오는 단어를 찾아서 순차 제시하고, all_text_cache의 후보를 다 소진했는데도 사용자가 계속 요구하는 경우, text_extract를 실행하여 all_text_cache를 업데이트하고 diff하여 달라진 부분만 매칭, 그리고 all_text_cache를 업데이트
// 3. prefix가 있는 경우, word_cache에서 채택빈도순으로 제시하고, 다 제시했는데도 사용자가 계속 요구하는 경우, all_text_cache를 업데이트하고 제시하지 않은 나머지 candidates들도 제시
//
// 사용자가 확정하면, word cache에 빈도 기록.

use std::cell::RefCell;
use std::collections::{HashMap, HashSet};

use dabbrev::{DabbrevState, WordCache};
use hwp_addon::debug::log;
use hwp_addon::export_hwp_addon;
use hwp_addon::hwp_user_action::{ActionMeta, HwpUserAction, ToolbarConfig, ToolbarTarget};
use hwp_addon::shortcut::{Modifiers, ShortcutKey};
use hwp_addon::text_extract::{HwpTextExt, TextSegment};
use hwp_core::hwp_obj::HwpObject;
use windows::Win32::UI::Input::KeyboardAndMouse::VK_OEM_2;

mod dabbrev;
mod ui_popup;

const TOOLBAR_DATA: &[u8] = include_bytes!("../toolbar.bmp");

const ACTION_DABBREV: &str = "DabbrevExpand";

static CONFIG: ToolbarConfig = ToolbarConfig {
    name: "dabbrev",
    serialize_path: "HwpDabbrev",
    bitmap_data: TOOLBAR_DATA,
    target: ToolbarTarget::Ribbon("자동완성"),
    ribbon_toolbox_index: -1,
};

// ── All Text Cache ──

pub(crate) struct AllTextCache {
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

fn extract_words(text: &str) -> impl Iterator<Item = &str> {
    text.split(|c: char| !is_word_char(c))
        .filter(|w| !w.is_empty())
}

fn is_word_char(c: char) -> bool {
    c.is_alphanumeric() || c == '_' || c == '.' || c == '-' || c == ':'
}

// ── Plugin ──

pub struct DabbrevPlugin {
    state: RefCell<Option<DabbrevState>>,
    /// 문서 path를 키로 사용. 저장 안 된 문서는 빈 문자열("")로 통합.
    /// HashMap::new()가 const fn이 아니므로 Option으로 감싸 lazy init.
    pub(crate) word_caches: RefCell<Option<HashMap<String, WordCache>>>,
    pub(crate) all_text_caches: RefCell<Option<HashMap<String, AllTextCache>>>,
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
        word_caches: RefCell::new(None),
        all_text_caches: RefCell::new(None),
    }
);
