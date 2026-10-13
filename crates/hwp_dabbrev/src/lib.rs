#![forbid(unsafe_code)]

// 액션마다 문서 전체에서 단어를 새로 추출하고, 완성 후보를 한 번에 표시한다.
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
        .map(strip_leading_nonword)
        .filter(|w| !w.is_empty())
}

pub(crate) fn is_word_char(c: char) -> bool {
    c.is_alphanumeric() || c == '_' || c == '-' || c == ':'
}

/// 단어/prefix는 영숫자에서 시작한다. 앞에 붙은 비단어 문자(연결용 구두점
/// `-`/`:`/`_` 등, 단어 내부에서만 의미 있는 문자)를 제거해 실제 단어
/// 시작점부터 반환한다. 예: `-bar`→`bar`. 내부 연결 문자는 보존(`a-b:c`).
pub(crate) fn strip_leading_nonword(w: &str) -> &str {
    w.trim_start_matches(|c: char| !c.is_alphanumeric())
}

// ── Plugin ──

pub struct DabbrevPlugin;

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
