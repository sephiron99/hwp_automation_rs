/// 각주·미주·주석 관련 액션
///
/// SDK 참고: ActionTable_2504.pdf § Note*, SaveFootnote
use crate::hwp_obj::HwpObject;

impl HwpObject {
    /// `NoteDelete` — 주석 지우기
    pub fn note_delete(&self) -> crate::error::Result<()> {
        self.h_action()?.run("NoteDelete")
    }

    /// `NoteModify` — 주석 고치기
    pub fn note_modify(&self) -> crate::error::Result<()> {
        self.h_action()?.run("NoteModify")
    }

    /// `NoteNumProperty` — 주석 번호 속성
    pub fn note_num_property(&self) -> crate::error::Result<()> {
        self.h_action()?.run("NoteNumProperty")
    }

    /// `NoteNoSuperscript` — 주석 번호 보통 (윗 첨자 사용 안함, ParameterSet: `SecDef`)
    pub fn note_no_superscript(&self) -> crate::error::Result<()> {
        self.h_action()?.run("NoteNoSuperscript")
    }

    /// `NoteSuperscript` — 주석 번호 작게 (윗 첨자, ParameterSet: `SecDef`)
    pub fn note_superscript(&self) -> crate::error::Result<()> {
        self.h_action()?.run("NoteSuperscript")
    }

    /// `NoteToNext` — 주석 다음으로 이동
    pub fn note_to_next(&self) -> crate::error::Result<()> {
        self.h_action()?.run("NoteToNext")
    }

    /// `NoteToPrev` — 주석 앞으로 이동
    pub fn note_to_prev(&self) -> crate::error::Result<()> {
        self.h_action()?.run("NoteToPrev")
    }

    /// `SaveFootnote` — 주석 저장 (ParameterSet: `SaveFootnote`)
    pub fn save_footnote(&self) -> crate::error::Result<()> {
        self.h_action()?.run("SaveFootnote")
    }
}
