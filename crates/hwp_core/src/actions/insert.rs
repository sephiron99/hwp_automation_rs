/// 삽입 관련 액션
///
/// SDK 참고: ActionTable_2504.pdf § Insert*, NewNumber*, OleCreateNew, VerticalText 등
///
/// `InsertText`는 [`crate::h_action`]의 `HwpObject::insert_text()`를 사용하십시오.
use crate::hwp_obj::HwpObject;

impl HwpObject {
    // ── 각주·미주 ──

    /// `InsertFootnote` — 각주를 삽입합니다.
    pub fn insert_footnote(&self) -> crate::error::Result<()> {
        self.h_action()?.run("InsertFootnote")
    }

    /// `InsertEndnote` — 미주를 삽입합니다.
    pub fn insert_endnote(&self) -> crate::error::Result<()> {
        self.h_action()?.run("InsertEndnote")
    }

    // ── 그림·파일 ──

    /// `InsertPicture` — 그림 삽입 대화상자 (ParameterSet: `InsertPicture`)
    pub fn insert_picture(&self) -> crate::error::Result<()> {
        self.h_action()?.run("InsertPicture")
    }

    /// `InsertFile` — 파일 삽입 대화상자 (ParameterSet: `InsertFile`)
    pub fn insert_file(&self) -> crate::error::Result<()> {
        self.h_action()?.run("InsertFile")
    }

    // ── 캡션 ──

    /// `InsertCaption` — 캡션 삽입 (ParameterSet: `CaptionDef`)
    pub fn insert_caption(&self) -> crate::error::Result<()> {
        self.h_action()?.run("InsertCaption")
    }

    /// `InsertCaptionDlg` — 캡션 삽입 대화상자 (ParameterSet: `CaptionDef`)
    pub fn insert_caption_dlg(&self) -> crate::error::Result<()> {
        self.h_action()?.run("InsertCaptionDlg")
    }

    /// `InsertFrameCaption` — 프레임 캡션 삽입
    pub fn insert_frame_caption(&self) -> crate::error::Result<()> {
        self.h_action()?.run("InsertFrameCaption")
    }

    // ── 필드·컨트롤 ──

    /// `InsertField` — 필드 삽입 (ParameterSet: `FieldCreate`)
    pub fn insert_field(&self) -> crate::error::Result<()> {
        self.h_action()?.run("InsertField")
    }

    /// `InsertCtrl` — 컨트롤 삽입 (ParameterSet: `CtrlCreate`)
    pub fn insert_ctrl(&self) -> crate::error::Result<()> {
        self.h_action()?.run("InsertCtrl")
    }

    // ── 특수 문자·자동 번호 ──

    /// `InsertSpecialChar` — 특수 문자 삽입 대화상자
    pub fn insert_special_char(&self) -> crate::error::Result<()> {
        self.h_action()?.run("InsertSpecialChar")
    }

    /// `InsertAutoNum` — 자동 번호 삽입 (ParameterSet: `AutoNum`)
    pub fn insert_auto_num(&self) -> crate::error::Result<()> {
        self.h_action()?.run("InsertAutoNum")
    }

    /// `InsertPageNum` — 쪽 번호 삽입 (ParameterSet: `PageNumPos`)
    pub fn insert_page_num(&self) -> crate::error::Result<()> {
        self.h_action()?.run("InsertPageNum")
    }

    /// `InsertLine` — 선 삽입
    pub fn insert_line(&self) -> crate::error::Result<()> {
        self.h_action()?.run("InsertLine")
    }

    // ── 번호 ──

    /// `NewNumber` — 새 번호로 시작 (ParameterSet: `AutoNum`)
    pub fn new_number(&self) -> crate::error::Result<()> {
        self.h_action()?.run("NewNumber")
    }

    /// `NewNumberModify` — 새 번호 고치기 (ParameterSet: `AutoNum`)
    pub fn new_number_modify(&self) -> crate::error::Result<()> {
        self.h_action()?.run("NewNumberModify")
    }

    /// `SelectPageNumShape` — 쪽 번호 모양 선택 (ParameterSet: `AutoNum`)
    pub fn select_page_num_shape(&self) -> crate::error::Result<()> {
        self.h_action()?.run("SelectPageNumShape")
    }

    // ── OLE ──

    /// `OleCreateNew` — OLE 개체 삽입 대화상자 (ParameterSet: `OleCreation`)
    pub fn ole_create_new(&self) -> crate::error::Result<()> {
        self.h_action()?.run("OleCreateNew")
    }

    // ── 기타 삽입 ──

    /// `VerticalText` — 세로쓰기 (ParameterSet: `TextVertical`)
    pub fn vertical_text(&self) -> crate::error::Result<()> {
        self.h_action()?.run("VerticalText")
    }

    /// `RecentCode` — 최근에 사용한 문자표 입력
    pub fn recent_code(&self) -> crate::error::Result<()> {
        self.h_action()?.run("RecentCode")
    }
}
