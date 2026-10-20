//! 문서 단위 편의 API — [`HwpObject::ext_doc`]
use crate::error::{HwpError, Result};
use crate::hwp_obj::HwpObject;
use crate::method::hwp_object::GetTextFileFormat;

use super::HwpmlCaret;

/// [`ExtDoc::hwpml_with_caret`]가 커서 자리에 잠깐 넣는 책갈피의 이름
const CARET_BOOKMARK: &str = "__hwp_core_caret__";

/// 문서 단위 편의 API 핸들. [`HwpObject::ext_doc`]로 얻는다.
pub struct ExtDoc<'a>(&'a HwpObject);

impl HwpObject {
    /// 문서 단위 편의 API 핸들을 돌려줍니다. [`ExtDoc`] 참고.
    pub fn ext_doc(&self) -> ExtDoc<'_> {
        ExtDoc(self)
    }
}

impl ExtDoc<'_> {
    /// 현재 문서를 hwpx(`HWPX`) 포맷으로 저장합니다. (`SaveAs(path, "HWPX", "")`)
    pub fn save_as_hwpx(&self, path: &str) -> Result<()> {
        self.0.save_as(path, "HWPX", "")
    }

    /// 현재 문서의 HWPML과, 그 안에서 커서가 있는 위치를 반환합니다.
    ///
    /// `GetPos`의 `(list, para, pos)`와 HWPML의 `CARETPOS`는 같은 번호를 담고 있지만,
    /// HWPML 본문 요소에는 그 번호에 대응하는 속성이 없습니다. 그래서 커서 자리에
    /// 책갈피를 넣고 `GetTextFile("HWPML2X")`을 받은 뒤, 실행 취소(`Undo`)로 책갈피를
    /// 되돌리고 HWPML을 한 번 더 받아, 두 HWPML의 본문이 책갈피 앞까지 같은 만큼을
    /// 커서 위치로 삼습니다.
    ///
    /// 문서를 잠깐 수정하므로 문서의 수정 상태와 다시 실행(`Redo`) 기록이 바뀝니다.
    /// `GetTextFile`을 두 번 부르므로 큰 문서에서는 느립니다.
    ///
    /// 선택 블록이 있으면 선택 방향과 상관없이 블록의 앞쪽 끝을 돌려주고, 선택이 풀리며
    /// 커서는 블록의 뒤쪽 끝으로 갑니다.
    pub fn hwpml_with_caret(&self) -> Result<HwpmlCaret> {
        let action = self.0.h_action()?;
        let pset = self.0.h_parameter_set()?;
        let bookmark = pset.h_book_mark()?;
        let hset = bookmark.h_set()?;
        action.get_default("Bookmark", &hset)?;
        bookmark.set_name(CARET_BOOKMARK)?;
        bookmark.set_type(0)?;
        bookmark.set_command(0)?;
        action.execute("Bookmark", &hset)?;

        // 책갈피는 HWPML을 얻지 못했어도 되돌린다.
        let marked = self.0.get_text_file(GetTextFileFormat::HwpMl2X, false);
        action.undo()?;
        let marked = marked?;
        let hwpml = self.0.get_text_file(GetTextFileFormat::HwpMl2X, false)?;

        let marker = format!("<BOOKMARK Name=\"{CARET_BOOKMARK}\"/>");
        let marker_at = marked.find(&marker).ok_or_else(|| {
            HwpError::ExecutionFailed("HWPML에서 커서 책갈피를 찾지 못했습니다.".to_string())
        })?;

        // 책갈피는 그 앞의 `CHAR`를 닫거나 뒤에 빈 `CHAR`를 만드는 등 주변 표기를 바꾸므로,
        // 책갈피만 빼지 않고 원래 문서와 책갈피 앞까지 같은 만큼을 센다. `HEAD`의
        // `CARETPOS`는 책갈피 때문에 달라지므로 `BODY`부터 비교한다.
        let body = |s: &str| s.find("<BODY>").unwrap_or(0);
        let before_marker = marked.get(body(&marked)..marker_at).unwrap_or("");
        let body_at = body(&hwpml);
        let same: usize = before_marker
            .chars()
            .zip(hwpml[body_at..].chars())
            .take_while(|(a, b)| a == b)
            .map(|(a, _)| a.len_utf8())
            .sum();
        Ok(HwpmlCaret {
            hwpml,
            offset: body_at + same,
        })
    }
}
