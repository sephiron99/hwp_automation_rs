/// 현재 커서 위치의 글자 모양 정보 (FormatSpec 캡처용)
///
/// SDK 참고: ParameterSetTable_2504.pdf § CharShape (13번)
/// `HAction.GetDefault("CharShape", hset)` + `HParameterSet.HCharShape` 로 조회합니다.
#[derive(Debug, Clone, Default)]
pub struct CharShape {
    /// 글자 크기 (HWPUNIT, 1pt = 2540 HWPUNIT)
    pub height: i32,
    /// 한글 글꼴 이름
    pub face_name_hangul: String,
    /// 영문 글꼴 이름 (FaceNameLatin)
    pub face_name_latin: String,
    /// 굵게 (0 = off, 1 = on)
    pub bold: u32,
    /// 기울임 (0 = off, 1 = on)
    pub italic: u32,
    /// 밑줄 종류 (0 = 없음, 1 = bottom, 2 = center, 3 = top)
    pub underline_type: u32,
    /// 취소선 종류 (0 = 없음, 1 = red single, 2 = red double, 3 = text single, 4 = text double)
    pub strike_out_type: u32,
    /// 글자 색 (COLORREF: 0x00BBGGRR)
    pub text_color: u32,
    /// 장평 — 한글 (50 ~ 200%)
    pub ratio_hangul: u32,
    /// 자간 — 한글 (−50 ~ 50%)
    pub spacing_hangul: i32,
}

// =========================================================================
// IHwpObject — CharShape 조회 메서드
// =========================================================================

impl crate::hwp_obj::HwpObject {
    /// 내부적으로 `HAction.GetDefault("CharShape", hset)`을 호출하여
    /// `HParameterSet.HCharShape`에서 각 필드를 읽습니다.
    ///
    /// 글자 모양 중 특정 항목이 selection 내에서 서로 다른 속성을 가지고 있으면 아예 아이템 자체가 존재하지 않는다.
    ///
    /// Selection이 없으면 현재 커서 위치의 글자 모양을 읽어 [`CharShape`]로 반환합니다.
    ///
    /// # 용도 (eat_hwp_raw)
    /// `FormatSpec` 캡처 — LLM이 삽입할 텍스트에 원본 글자 모양을 재적용하기 위해 사용합니다.
    pub fn current_char_shape(&self) -> crate::error::Result<CharShape> {
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
}
