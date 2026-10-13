//! 편의 API 층 — SDK method와 Action을 조합한 고수준 함수
//!
//! SDK에 1:1로 대응하지 않는, 이 프로젝트가 자체 설계한 API다. 아래 층
//! ([`crate::pset`], [`crate::method`], [`crate::actions`])의 공개 API만으로 구현한다.
//!
//! ```ignore
//! use hwp_core::ext::HwpExt;
//!
//! hwp.insert_text("Hello")?;
//! ```
pub mod char_shape;
pub mod hwpml_caret;
pub mod para_shape;

pub use char_shape::CharShape;
pub use hwpml_caret::HwpmlCaret;
pub use para_shape::ParaShape;

use crate::error::{HwpError, Result};
use crate::hwp_obj::HwpObject;
use crate::method::hwp_object::GetTextFileFormat;

/// [`HwpExt::hwpml_with_caret`]가 커서 자리에 잠깐 넣는 책갈피의 이름
const CARET_BOOKMARK: &str = "__hwp_core_caret__";

/// [`HwpObject`]에 편의 메서드를 추가하는 extension trait.
pub trait HwpExt {
    /// 현재 커서 위치에 텍스트를 삽입합니다.
    ///
    /// SDK 액션 ID: `"InsertText"`, 파라미터: `HInsertText.Text`
    fn insert_text(&self, text: &str) -> Result<()>;

    /// 현재 문서를 hwpx(`HWPX`) 포맷으로 저장합니다. (`SaveAs(path, "HWPX", "")`)
    fn save_as_hwpx(&self, path: &str) -> Result<()>;

    /// 내부적으로 `HAction.GetDefault("CharShape", hset)`을 호출하여
    /// `HParameterSet.HCharShape`에서 각 필드를 읽습니다.
    ///
    /// 글자 모양 중 특정 항목이 selection 내에서 서로 다른 속성을 가지고 있으면 아예 아이템 자체가 존재하지 않는다.
    ///
    /// Selection이 없으면 현재 커서 위치의 글자 모양을 읽어 [`CharShape`]로 반환합니다.
    ///
    /// # 용도 (eat_hwp_raw)
    /// `FormatSpec` 캡처 — LLM이 삽입할 텍스트에 원본 글자 모양을 재적용하기 위해 사용합니다.
    fn current_char_shape(&self) -> Result<CharShape>;

    /// 현재 커서 위치의 문단 모양을 읽어 [`ParaShape`]로 반환합니다.
    ///
    /// 내부적으로 `HAction.GetDefault("ParaShape", hset)`을 호출하여
    /// `HParameterSet.HParaShape`에서 각 필드를 읽습니다.
    ///
    /// # 용도 (eat_hwp_raw)
    /// `FormatSpec` 캡처 — 삽입 후 원본 문단 모양을 재적용하거나
    /// 열 폭(column width) 추정에 활용합니다.
    fn current_para_shape(&self) -> Result<ParaShape>;

    /// 현재 문서의 HWPML과, 그 안에서 커서가 있는 위치를 반환합니다.
    ///
    /// `GetPos`의 `(list, para, pos)`와 HWPML의 `CARETPOS`는 같은 번호를 담고 있지만,
    /// HWPML 본문 요소에는 그 번호에 대응하는 속성이 없습니다. 그래서 커서 자리에
    /// 책갈피를 넣고 `GetTextFile("HWPML2X")`을 받은 뒤, 실행 취소(`Undo`)로 책갈피를
    /// 되돌리고 HWPML에서 그 책갈피가 있던 자리를 찾습니다.
    ///
    /// 문서를 잠깐 수정하므로 문서의 수정 상태와 다시 실행(`Redo`) 기록이 바뀝니다.
    /// 선택 블록이 없는 상태에서 호출합니다.
    fn hwpml_with_caret(&self) -> Result<HwpmlCaret>;
}

impl HwpExt for HwpObject {
    fn insert_text(&self, text: &str) -> Result<()> {
        let action = self.h_action()?;
        let pset = self.h_parameter_set()?;
        let insert_text = pset.h_insert_text()?;
        let hset = insert_text.h_set()?;
        action.get_default("InsertText", &hset)?;
        insert_text.set_text(text)?;
        action.execute("InsertText", &hset)?;
        Ok(())
    }

    fn save_as_hwpx(&self, path: &str) -> Result<()> {
        self.save_as(path, "HWPX", "")
    }

    fn current_char_shape(&self) -> Result<CharShape> {
        let action = self.h_action()?;
        let pset = self.h_parameter_set()?;
        let cs = pset.h_char_shape()?;
        let hset = cs.h_set()?;
        action.get_default("CharShape", &hset)?;

        Ok(CharShape {
            height: cs.get("Height").unwrap_or(0),
            face_name_hangul: cs.get("FaceNameHangul").unwrap_or_default(),
            face_name_latin: cs.get("FaceNameLatin").unwrap_or_default(),
            bold: cs.get("Bold").unwrap_or(0),
            italic: cs.get("Italic").unwrap_or(0),
            underline_type: cs.get("UnderlineType").unwrap_or(0),
            strike_out_type: cs.get("StrikeOutType").unwrap_or(0),
            text_color: cs.get("TextColor").unwrap_or(0),
            ratio_hangul: cs.get("RatioHangul").unwrap_or(100),
            spacing_hangul: cs.get("SpacingHangul").unwrap_or(0),
        })
    }

    fn current_para_shape(&self) -> Result<ParaShape> {
        let action = self.h_action()?;
        let pset = self.h_parameter_set()?;
        let ps = pset.h_para_shape()?;
        let hset = ps.h_set()?;
        action.get_default("ParaShape", &hset)?;

        Ok(ParaShape {
            align_type: ps.get("AlignType").unwrap_or(0),
            line_spacing_type: ps.get("LineSpacingType").unwrap_or(0),
            line_spacing: ps.get("LineSpacing").unwrap_or(160),
            left_margin: ps.get("LeftMargin").unwrap_or(0),
            right_margin: ps.get("RightMargin").unwrap_or(0),
            indentation: ps.get("Indentation").unwrap_or(0),
            prev_spacing: ps.get("PrevSpacing").unwrap_or(0),
            next_spacing: ps.get("NextSpacing").unwrap_or(0),
        })
    }

    fn hwpml_with_caret(&self) -> Result<HwpmlCaret> {
        let action = self.h_action()?;
        let pset = self.h_parameter_set()?;
        let bookmark = pset.h_book_mark()?;
        let hset = bookmark.h_set()?;
        action.get_default("Bookmark", &hset)?;
        bookmark.set_name(CARET_BOOKMARK)?;
        bookmark.set_type(0)?;
        bookmark.set_command(0)?;
        action.execute("Bookmark", &hset)?;

        // 책갈피는 HWPML을 얻지 못했어도 되돌린다.
        let hwpml = self.get_text_file(GetTextFileFormat::HwpMl2X, false);
        self.act().undo()?;
        let mut hwpml = hwpml?;

        let marker = format!("<BOOKMARK Name=\"{CARET_BOOKMARK}\"/>");
        let mut offset = hwpml.find(&marker).ok_or_else(|| {
            HwpError::ExecutionFailed("HWPML에서 커서 책갈피를 찾지 못했습니다.".to_string())
        })?;
        hwpml.replace_range(offset..offset + marker.len(), "");

        // 글자 사이에 넣은 책갈피는 `CHAR`를 둘로 나누므로, 나뉜 것을 다시 잇는다.
        const SPLIT: &str = "</CHAR><CHAR>";
        let close = "</CHAR>".len();
        if let Some(start) = offset.checked_sub(close)
            && hwpml.get(start..).is_some_and(|s| s.starts_with(SPLIT))
        {
            hwpml.replace_range(start..start + SPLIT.len(), "");
            offset = start;
        }
        Ok(HwpmlCaret { hwpml, offset })
    }
}
