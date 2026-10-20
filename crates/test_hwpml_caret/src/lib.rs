//! `ExtDoc::hwpml_with_caret` 실환경 검증용 애드온 DLL
//!
//! 툴바의 "HWPML 커서" 버튼을 누르면 현재 커서 자리에서 `hwpml_with_caret`을 실행하고
//! 결과를 복사할 수 있는 창으로 보여 준다. 확인하는 것은 `hwp_실전_노트.md` 11절의 미확인 가정이다.
use hwp_addon::debug::msgbox_err;
use hwp_addon::hwp_user_action::{ActionMeta, HwpUserAction, ToolbarConfig, ToolbarTarget};
use hwp_addon::method::hwp_object::GetTextFileFormat;
use hwp_addon::{HwpObject, export_hwp_addon};

mod result_window;

const TOOLBAR_DATA: &[u8] = include_bytes!("../toolbar.bmp");

const MY_ACTION: &str = "TestHwpmlCaret";

/// 커서 앞뒤로 보여 줄 HWPML 글자 수
const CONTEXT_CHARS: usize = 120;

static CONFIG: ToolbarConfig = ToolbarConfig {
    name: "addon",
    serialize_path: "TestHwpmlCaret",
    bitmap_data: TOOLBAR_DATA,
    target: ToolbarTarget::Both("추가 기능"),
    ribbon_toolbox_index: -1,
};

pub struct TestHwpmlCaretPlugin;

impl HwpUserAction for TestHwpmlCaretPlugin {
    fn toolbar_config(&self) -> Option<&ToolbarConfig> {
        Some(&CONFIG)
    }

    fn actions(&self) -> &'static [ActionMeta] {
        static ACTIONS: [ActionMeta; 1] = [ActionMeta {
            name: MY_ACTION,
            label: "HWPML 커서",
            image_index: 0,
            shortcut: None,
        }];
        &ACTIONS
    }

    fn on_load(&self, hwp: &HwpObject) -> hwp_addon::Result<bool> {
        match self.setup_toolbar(hwp) {
            Ok(v) => Ok(v),
            Err(e) => {
                msgbox_err("on_load", &format!("{e}"));
                Err(e)
            }
        }
    }

    fn do_action(&self, action_name: &str, hwp: &HwpObject) -> hwp_addon::Result<bool> {
        match action_name {
            MY_ACTION => {
                let text = report(hwp).unwrap_or_else(|e| format!("실패: {e}"));
                result_window::show(MY_ACTION, &text);
                Ok(true)
            }
            _ => Ok(false),
        }
    }
}

/// `hwpml_with_caret`을 실행하고, 호출 전후의 커서와 문서를 비교한 결과를 글로 만든다.
fn report(hwp: &HwpObject) -> hwp_addon::Result<String> {
    let pos_before = hwp.get_pos()?;
    let hwpml_before = hwp.get_text_file(GetTextFileFormat::HwpMl2X, false)?;

    let doc = hwp.ext_doc().hwpml_with_caret()?;

    let pos_after = hwp.get_pos()?;
    let hwpml_after = hwp.get_text_file(GetTextFileFormat::HwpMl2X, false)?;

    let (head, tail) = doc.hwpml.split_at(doc.offset);
    let head_start = head
        .char_indices()
        .rev()
        .nth(CONTEXT_CHARS - 1)
        .map_or(0, |(i, _)| i);
    let tail_end = tail
        .char_indices()
        .nth(CONTEXT_CHARS)
        .map_or(tail.len(), |(i, _)| i);

    let diff = first_diff(&hwpml_before, &doc.hwpml);

    Ok(format!(
        "GetPos 전: {pos_before:?}\n\
         GetPos 후: {pos_after:?} ({})\n\n\
         호출 후 문서가 호출 전과 같음(Undo로 복원): {}\n\
         반환한 HWPML이 호출 전 문서와 같음: {}{diff}\n\n\
         offset: {} / {} 바이트\n\n\
         커서 앞:\n{}\n\n\
         커서 뒤:\n{}",
        verdict(pos_before == pos_after),
        verdict(hwpml_before == hwpml_after),
        verdict(hwpml_before == doc.hwpml),
        doc.offset,
        doc.hwpml.len(),
        &head[head_start..],
        &tail[..tail_end],
    ))
}

/// 두 문자열이 처음 달라지는 곳부터 각각 [`CONTEXT_CHARS`]자를 보여 준다. 같으면 빈 문자열.
fn first_diff(before: &str, returned: &str) -> String {
    let same: usize = before
        .chars()
        .zip(returned.chars())
        .take_while(|(a, b)| a == b)
        .map(|(a, _)| a.len_utf8())
        .sum();
    if same == before.len() && same == returned.len() {
        return String::new();
    }
    let excerpt = |s: &str| s[same..].chars().take(CONTEXT_CHARS).collect::<String>();
    format!(
        "\n  처음 달라지는 곳({same}바이트)\n  호출 전: {}\n  반환값: {}",
        excerpt(before),
        excerpt(returned),
    )
}

fn verdict(ok: bool) -> &'static str {
    if ok { "일치" } else { "불일치" }
}

export_hwp_addon!(TestHwpmlCaretPlugin, TestHwpmlCaretPlugin);
