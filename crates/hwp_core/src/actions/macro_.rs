/// 매크로·빠른 교정·빠른 책갈피 관련 액션
///
/// SDK 참고: ActionTable_2504.pdf § ScrMacro*, QuickCorrect*, QuickMark*, QuickCommand
use crate::hwp_obj::HwpObject;

impl HwpObject {
    // ── 스크립트 매크로 ──

    /// `ScrMacroDefine` — 매크로 정의 대화상자를 열고, 설정이 끝나면 매크로를 기록합니다.
    ///
    /// (ParameterSet: `ScriptMacro`)
    pub fn scr_macro_define(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ScrMacroDefine")
    }

    /// `ScrMacroPause` — 매크로 기록 일시정지/재시작
    pub fn scr_macro_pause(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ScrMacroPause")
    }

    /// `ScrMacroStop` — 매크로 기록 중지
    pub fn scr_macro_stop(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ScrMacroStop")
    }

    /// `ScrMacroRepeatDlg` — 스크립트 매크로 실행 대화상자 (ParameterSet: `ScriptMacro`)
    pub fn scr_macro_repeat_dlg(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ScrMacroRepeatDlg")
    }

    /// `ScrMacroPlay1` — #1 매크로 실행 (Alt+Shift+1)
    pub fn scr_macro_play1(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ScrMacroPlay1")
    }

    /// `ScrMacroPlay2` — #2 매크로 실행 (Alt+Shift+2)
    pub fn scr_macro_play2(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ScrMacroPlay2")
    }

    /// `ScrMacroPlay3` — #3 매크로 실행 (Alt+Shift+3)
    pub fn scr_macro_play3(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ScrMacroPlay3")
    }

    /// `ScrMacroPlay4` — #4 매크로 실행 (Alt+Shift+4)
    pub fn scr_macro_play4(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ScrMacroPlay4")
    }

    /// `ScrMacroPlay5` — #5 매크로 실행 (Alt+Shift+5)
    pub fn scr_macro_play5(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ScrMacroPlay5")
    }

    /// `ScrMacroPlay6` — #6 매크로 실행 (Alt+Shift+6)
    pub fn scr_macro_play6(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ScrMacroPlay6")
    }

    /// `ScrMacroPlay7` — #7 매크로 실행 (Alt+Shift+7)
    pub fn scr_macro_play7(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ScrMacroPlay7")
    }

    /// `ScrMacroPlay8` — #8 매크로 실행 (Alt+Shift+8)
    pub fn scr_macro_play8(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ScrMacroPlay8")
    }

    /// `ScrMacroPlay9` — #9 매크로 실행 (Alt+Shift+9)
    pub fn scr_macro_play9(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ScrMacroPlay9")
    }

    /// `ScrMacroPlay10` — #10 매크로 실행 (Alt+Shift+0)
    pub fn scr_macro_play10(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ScrMacroPlay10")
    }

    /// `ScrMacroPlay11` — #11 매크로 실행
    pub fn scr_macro_play11(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ScrMacroPlay11")
    }

    // ── 빠른 교정 ──

    /// `QuickCommandRun` — 입력 자동 명령 동작 (`QuickCommand Run`)
    pub fn quick_command_run(&self) -> crate::error::Result<()> {
        self.h_action()?.run("QuickCommand Run")
    }

    /// `QuickCorrectEdit` — 빠른 교정 내용 편집 (ParameterSet: `QCorrect`)
    pub fn quick_correct_edit(&self) -> crate::error::Result<()> {
        self.h_action()?.run("QuickCorrect Edit")
    }

    /// `QuickCorrectRun` — 빠른 교정 내용 편집 실행 (`QuickCorrect Run`)
    pub fn quick_correct_run(&self) -> crate::error::Result<()> {
        self.h_action()?.run("QuickCorrect Run")
    }

    /// `QuickCorrectSound` — 빠른 교정 메뉴에서 효과음 On/Off (`QuickCorrect Sound`)
    pub fn quick_correct_sound(&self) -> crate::error::Result<()> {
        self.h_action()?.run("QuickCorrect Sound")
    }

    /// `QuickCorrect` — 빠른 교정 (실질적인 동작 Action)
    pub fn quick_correct(&self) -> crate::error::Result<()> {
        self.h_action()?.run("QuickCorrect")
    }

    // ── 빠른 책갈피 ──

    /// `QuickMarkInsert0` — 쉬운 책갈피 삽입 (0번)
    pub fn quick_mark_insert0(&self) -> crate::error::Result<()> {
        self.h_action()?.run("QuickMarkInsert0")
    }

    /// `QuickMarkInsert1` — 쉬운 책갈피 삽입 (1번)
    pub fn quick_mark_insert1(&self) -> crate::error::Result<()> {
        self.h_action()?.run("QuickMarkInsert1")
    }

    /// `QuickMarkInsert2` — 쉬운 책갈피 삽입 (2번)
    pub fn quick_mark_insert2(&self) -> crate::error::Result<()> {
        self.h_action()?.run("QuickMarkInsert2")
    }

    /// `QuickMarkInsert3` — 쉬운 책갈피 삽입 (3번)
    pub fn quick_mark_insert3(&self) -> crate::error::Result<()> {
        self.h_action()?.run("QuickMarkInsert3")
    }

    /// `QuickMarkInsert4` — 쉬운 책갈피 삽입 (4번)
    pub fn quick_mark_insert4(&self) -> crate::error::Result<()> {
        self.h_action()?.run("QuickMarkInsert4")
    }

    /// `QuickMarkInsert5` — 쉬운 책갈피 삽입 (5번)
    pub fn quick_mark_insert5(&self) -> crate::error::Result<()> {
        self.h_action()?.run("QuickMarkInsert5")
    }

    /// `QuickMarkInsert6` — 쉬운 책갈피 삽입 (6번)
    pub fn quick_mark_insert6(&self) -> crate::error::Result<()> {
        self.h_action()?.run("QuickMarkInsert6")
    }

    /// `QuickMarkInsert7` — 쉬운 책갈피 삽입 (7번)
    pub fn quick_mark_insert7(&self) -> crate::error::Result<()> {
        self.h_action()?.run("QuickMarkInsert7")
    }

    /// `QuickMarkInsert8` — 쉬운 책갈피 삽입 (8번)
    pub fn quick_mark_insert8(&self) -> crate::error::Result<()> {
        self.h_action()?.run("QuickMarkInsert8")
    }

    /// `QuickMarkInsert9` — 쉬운 책갈피 삽입 (9번)
    pub fn quick_mark_insert9(&self) -> crate::error::Result<()> {
        self.h_action()?.run("QuickMarkInsert9")
    }

    /// `QuickMarkMove0` — 쉬운 책갈피 이동 (0번)
    pub fn quick_mark_move0(&self) -> crate::error::Result<()> {
        self.h_action()?.run("QuickMarkMove0")
    }

    /// `QuickMarkMove1` — 쉬운 책갈피 이동 (1번)
    pub fn quick_mark_move1(&self) -> crate::error::Result<()> {
        self.h_action()?.run("QuickMarkMove1")
    }

    /// `QuickMarkMove2` — 쉬운 책갈피 이동 (2번)
    pub fn quick_mark_move2(&self) -> crate::error::Result<()> {
        self.h_action()?.run("QuickMarkMove2")
    }

    /// `QuickMarkMove3` — 쉬운 책갈피 이동 (3번)
    pub fn quick_mark_move3(&self) -> crate::error::Result<()> {
        self.h_action()?.run("QuickMarkMove3")
    }

    /// `QuickMarkMove4` — 쉬운 책갈피 이동 (4번)
    pub fn quick_mark_move4(&self) -> crate::error::Result<()> {
        self.h_action()?.run("QuickMarkMove4")
    }

    /// `QuickMarkMove5` — 쉬운 책갈피 이동 (5번)
    pub fn quick_mark_move5(&self) -> crate::error::Result<()> {
        self.h_action()?.run("QuickMarkMove5")
    }

    /// `QuickMarkMove6` — 쉬운 책갈피 이동 (6번)
    pub fn quick_mark_move6(&self) -> crate::error::Result<()> {
        self.h_action()?.run("QuickMarkMove6")
    }

    /// `QuickMarkMove7` — 쉬운 책갈피 이동 (7번)
    pub fn quick_mark_move7(&self) -> crate::error::Result<()> {
        self.h_action()?.run("QuickMarkMove7")
    }

    /// `QuickMarkMove8` — 쉬운 책갈피 이동 (8번)
    pub fn quick_mark_move8(&self) -> crate::error::Result<()> {
        self.h_action()?.run("QuickMarkMove8")
    }

    /// `QuickMarkMove9` — 쉬운 책갈피 이동 (9번)
    pub fn quick_mark_move9(&self) -> crate::error::Result<()> {
        self.h_action()?.run("QuickMarkMove9")
    }
}
