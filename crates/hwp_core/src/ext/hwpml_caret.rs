/// HWPML 문서와 그 안에서의 커서 위치
///
/// [`HwpExt::hwpml_with_caret`](super::HwpExt::hwpml_with_caret)가 반환합니다.
#[derive(Debug, Clone, Default)]
pub struct HwpmlCaret {
    /// 현재 문서의 HWPML2X 문자열. `GetTextFile("HWPML2X")`이 돌려주는 그대로다.
    pub hwpml: String,
    /// `hwpml` 안에서 커서가 있는 바이트 위치. `<TEXT>` 안을 가리키며, 커서가 글자
    /// 사이에 있으면 `<CHAR>`의 글 안이고 그렇지 않으면 요소 경계다. 빈 문단은
    /// `TEXT`가 `<TEXT .../>`로 닫혀 있어 `/>` 바로 앞을 가리킨다.
    pub offset: usize,
}
