//! 글자·문단 모양 편의 API — [`HwpObject::ext_shape`]
use crate::error::Result;
use crate::hwp_obj::HwpObject;

use super::{CharShape, ParaShape};

/// 글자·문단 모양 편의 API 핸들. [`HwpObject::ext_shape`]로 얻는다.
pub struct ExtShape<'a>(&'a HwpObject);

impl HwpObject {
    /// 글자·문단 모양 편의 API 핸들을 돌려줍니다. [`ExtShape`] 참고.
    pub fn ext_shape(&self) -> ExtShape<'_> {
        ExtShape(self)
    }
}

impl ExtShape<'_> {
    /// 내부적으로 `HAction.GetDefault("CharShape", hset)`을 호출하여
    /// `HParameterSet.HCharShape`에서 각 필드를 읽습니다.
    ///
    /// 글자 모양 중 특정 항목이 selection 내에서 서로 다른 속성을 가지고 있으면 아예 아이템 자체가 존재하지 않는다.
    ///
    /// Selection이 없으면 현재 커서 위치의 글자 모양을 읽어 [`CharShape`]로 반환합니다.
    ///
    /// # 용도 (eat_hwp_raw)
    /// `FormatSpec` 캡처 — LLM이 삽입할 텍스트에 원본 글자 모양을 재적용하기 위해 사용합니다.
    pub fn current_char(&self) -> Result<CharShape> {
        let action = self.0.h_action()?;
        let pset = self.0.h_parameter_set()?;
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

    /// 현재 커서 위치의 문단 모양을 읽어 [`ParaShape`]로 반환합니다.
    ///
    /// 내부적으로 `HAction.GetDefault("ParaShape", hset)`을 호출하여
    /// `HParameterSet.HParaShape`에서 각 필드를 읽습니다.
    ///
    /// # 용도 (eat_hwp_raw)
    /// `FormatSpec` 캡처 — 삽입 후 원본 문단 모양을 재적용하거나
    /// 열 폭(column width) 추정에 활용합니다.
    pub fn current_para(&self) -> Result<ParaShape> {
        let action = self.0.h_action()?;
        let pset = self.0.h_parameter_set()?;
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
