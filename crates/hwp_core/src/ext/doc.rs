//! 문서 단위 편의 API — [`HwpObject::ext_doc`]
use crate::error::{HwpError, Result};
use crate::hwp_obj::HwpObject;
use crate::method::hwp_object::GetTextFileFormat;

use crate::method::h_action::HAction;

use super::{HwpmlCaret, HwpxCaret};

/// [`ExtDoc::hwpml_with_caret`]가 커서 자리에 잠깐 넣는 책갈피의 이름
const CARET_BOOKMARK: &str = "__hwp_core_caret__";

/// [`ExtDoc::hwpx_with_caret`]가 넣는 책갈피 이름의 접두어
const HWPX_CARET_PREFIX: &str = "_CARET_";

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
        let action = self.insert_bookmark(CARET_BOOKMARK)?;

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

    /// 커서 자리에 책갈피를 넣은 현재 문서를 사용자 임시 디렉터리에 HWPX로 저장합니다.
    ///
    /// 책갈피 이름은 `_CARET_` 뒤에 한글·영문 대소문자·숫자 5글자를 무작위로 붙여, 문서에
    /// 이미 있는 책갈피와 겹치지 않게 합니다. 저장 파일은 임시 디렉터리의
    /// `<책갈피 이름>.hwpx`입니다. 저장한 뒤 실행 취소(`Undo`)로 책갈피를 되돌리므로 열린
    /// 문서에는 책갈피가 남지 않고, 저장한 HWPX에만 남습니다.
    ///
    /// 저장한 파일은 지우지 않습니다. 다 쓴 뒤 호출자가 지워야 합니다.
    ///
    /// `SaveAs`로 저장해도 열린 문서의 경로는 바뀌지 않고, 커서 위치도 그대로입니다(한글
    /// 2018에서 확인). 다만 책갈피를 넣고 되돌리므로 다시 실행(`Redo`) 기록이 바뀌고, 수정
    /// 상태도 바뀔 수 있습니다. 선택 블록이 있을 때의 동작은
    /// [`hwpml_with_caret`](Self::hwpml_with_caret)와 같을 것으로 보지만 확인하지 않았습니다.
    pub fn hwpx_with_caret(&self) -> Result<HwpxCaret> {
        let bookmark_name = random_caret_bookmark();
        let hwpx = std::env::temp_dir().join(format!("{bookmark_name}.hwpx"));
        let path = hwpx.to_str().ok_or_else(|| {
            HwpError::ExecutionFailed(format!(
                "임시 파일 경로를 문자열로 바꿀 수 없습니다: {}",
                hwpx.display()
            ))
        })?;

        let action = self.insert_bookmark(&bookmark_name)?;
        // 책갈피는 저장에 실패해도 되돌린다. `lock:false`는 저장한 파일을 한글이 잠그지
        // 않게 해 호출자가 읽고 지울 수 있게 한다.
        let saved = self.0.save_as(path, "HWPX", "lock:false");
        action.undo()?;
        saved?;

        Ok(HwpxCaret {
            hwpx,
            bookmark_name,
        })
    }

    /// 커서 자리에 `name` 책갈피를 넣고, 되돌릴 때 쓸 `HAction`을 돌려준다.
    fn insert_bookmark(&self, name: &str) -> Result<HAction> {
        let action = self.0.h_action()?;
        let pset = self.0.h_parameter_set()?;
        let bookmark = pset.h_book_mark()?;
        let hset = bookmark.h_set()?;
        action.get_default("Bookmark", &hset)?;
        bookmark.set_name(name)?;
        bookmark.set_type(0)?;
        bookmark.set_command(0)?;
        action.execute("Bookmark", &hset)?;
        Ok(action)
    }
}

/// `_CARET_` 뒤에 한글 음절·영문 대소문자·숫자 가운데 5글자를 무작위로 붙인 책갈피 이름
///
/// 난수는 표준 라이브러리 `RandomState`의 프로세스별 무작위 키로 만든다. 암호용이 아니라
/// 문서의 다른 책갈피와 겹치지 않게 하는 용도다.
fn random_caret_bookmark() -> String {
    use std::hash::{BuildHasher, Hasher};

    const ASCII: &[u8] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
    const HANGUL_FIRST: u32 = 0xAC00; // 가
    const HANGUL_COUNT: u64 = 11172; // 가..=힣
    let total = HANGUL_COUNT + ASCII.len() as u64;

    let state = std::collections::hash_map::RandomState::new();
    let now = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map_or(0, |d| d.as_nanos());
    let mut name = String::from(HWPX_CARET_PREFIX);
    for i in 0..5u8 {
        let mut h = state.build_hasher();
        h.write_u128(now);
        h.write_u8(i);
        let n = h.finish() % total;
        let c = if n < HANGUL_COUNT {
            char::from_u32(HANGUL_FIRST + n as u32).expect("한글 음절 범위")
        } else {
            ASCII[(n - HANGUL_COUNT) as usize] as char
        };
        name.push(c);
    }
    name
}
