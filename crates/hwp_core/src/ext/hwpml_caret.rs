use std::path::{Path, PathBuf};

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
///
/// 임시 파일의 수명은 이 값의 수명과 같습니다. 값이 버려지면(drop) 파일을 지웁니다.
/// 파일을 남기려면 [`keep`](Self::keep)을, 지우기 실패를 알아야 하면 [`close`](Self::close)를
/// 씁니다. 파일이 다른 프로그램에 열려 있어 drop 때 지우지 못하면 조용히 남습니다.
/// 프로세스가 비정상 종료하면 drop이 실행되지 않아 파일이 남습니다.
#[derive(Debug)]
pub struct HwpxCaret {
    /// 저장한 `.hwpx` 파일. `keep`이나 `close`로 넘겨준 뒤에는 비어 있다.
    hwpx: PathBuf,
    bookmark_name: String,
}

impl HwpxCaret {
    pub(crate) fn new(hwpx: PathBuf, bookmark_name: String) -> Self {
        Self {
            hwpx,
            bookmark_name,
        }
    }

    /// 현재 문서를 저장한 사용자 임시 디렉터리의 `.hwpx` 파일 경로
    pub fn path(&self) -> &Path {
        &self.hwpx
    }

    /// 커서 자리를 알리는 책갈피 이름. `_CARET_` 뒤에 한글·영문 대소문자·숫자 5글자가 붙는다.
    pub fn bookmark_name(&self) -> &str {
        &self.bookmark_name
    }

    /// 파일을 지우지 않고 경로를 넘겨받습니다. 이후 삭제는 호출자가 합니다.
    pub fn keep(mut self) -> PathBuf {
        std::mem::take(&mut self.hwpx)
    }

    /// 파일을 지금 지우고, 실패하면 에러를 돌려줍니다.
    pub fn close(mut self) -> std::io::Result<()> {
        std::fs::remove_file(std::mem::take(&mut self.hwpx))
    }
}

impl Drop for HwpxCaret {
    fn drop(&mut self) {
        if !self.hwpx.as_os_str().is_empty() {
            // drop에서는 에러를 돌려줄 수 없다. 실패를 알아야 하면 `close`를 쓴다.
            let _ = std::fs::remove_file(&self.hwpx);
        }
    }
}
