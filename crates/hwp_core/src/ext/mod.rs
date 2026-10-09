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
pub mod para_shape;

pub use char_shape::CharShape;
pub use para_shape::ParaShape;

use crate::error::Result;
use crate::hwp_obj::HwpObject;

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
}
