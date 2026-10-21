/// HWPML 문서와 그 안에서의 커서 위치
///
/// [`ExtDoc::hwpml_with_caret`](super::ExtDoc::hwpml_with_caret)가 반환합니다.
#[derive(Debug, Clone, Default)]
pub struct HwpmlCaret {
    /// 현재 문서의 HWPML2X 문자열. `GetTextFile("HWPML2X")`이 돌려주는 그대로다.
    pub hwpml: String,
    /// `hwpml` 안에서 커서가 있는 바이트 위치. `<TEXT>` 안을 가리키며, 커서가 글자
    /// 사이에 있으면 `<CHAR>`의 글 안이고 그렇지 않으면 요소 경계다. 빈 문단은
    /// `TEXT`가 `<TEXT .../>`로 닫혀 있어 `/>` 바로 앞을 가리킨다.
    pub offset: usize,
}

/// 커서 자리에 책갈피를 넣어 저장한 HWPX 파일
///
/// [`ExtDoc::hwpx_with_caret`](super::ExtDoc::hwpx_with_caret)가 반환합니다.
#[derive(Debug, Clone, Default)]
pub struct HwpxCaret {
    /// 현재 문서를 저장한 사용자 임시 디렉터리의 `.hwpx` 파일. 다 쓴 뒤 호출자가 지운다.
    pub hwpx: std::path::PathBuf,
    /// 커서 자리를 알리는 책갈피 이름. `_CARET_` 뒤에 한글·영문 대소문자·숫자 5글자가 붙는다.
    pub bookmark_name: String,
}
