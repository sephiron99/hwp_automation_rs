/// HWPML 문서와 그 안에서의 커서 위치
///
/// [`HwpExt::hwpml_with_caret`](super::HwpExt::hwpml_with_caret)가 반환합니다.
#[derive(Debug, Clone, Default)]
pub struct HwpmlCaret {
    /// 현재 문서의 HWPML2X 문자열 (커서 표식은 빠져 있음)
    ///
    /// 책갈피를 넣은 직후에 받은 것이라 `HEAD`의 `CARETPOS`는 원래 커서와 다를 수 있다.
    pub hwpml: String,
    /// `hwpml` 안에서 커서가 있는 바이트 위치. `<TEXT>` 안을 가리키며, 커서가 글자
    /// 사이에 있으면 `<CHAR>`의 글 안이고 그렇지 않으면 요소 경계다.
    pub offset: usize,
}
