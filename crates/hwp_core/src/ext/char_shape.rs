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
