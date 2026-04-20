/// 찾기·바꾸기·검색 관련 액션
///
/// SDK 참고: ActionTable_2504.pdf § FindDlg, BackwardFind, ForwardFind, ReplaceDlg 등
///
/// 파라미터가 있는 액션(`FindReplace*`)은 대화상자 없이 호출하면 기본값으로 동작합니다.
/// 파라미터를 설정하여 프로그래밍 방식으로 찾으려면 `HAction.GetDefault + Execute` 패턴을 사용하십시오.
use crate::hwp_obj::HwpObject;

impl HwpObject {
    /// `FindDlg` — 찾기 대화상자를 표시합니다.
    pub fn find_dlg(&self) -> crate::error::Result<()> {
        self.h_action()?.run("FindDlg")
    }

    /// `FindAll` — 모두 찾기 (ParameterSet: `FindReplace*`)
    pub fn find_all(&self) -> crate::error::Result<()> {
        self.h_action()?.run("FindAll")
    }

    /// `BackwardFind` — 뒤로 찾기 (ParameterSet: `FindReplace*`)
    pub fn backward_find(&self) -> crate::error::Result<()> {
        self.h_action()?.run("BackwardFind")
    }

    /// `ForwardFind` — 앞으로 찾기 (ParameterSet: `FindReplace*`)
    pub fn forward_find(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ForwardFind")
    }

    /// `RepeatFind` — 다시 찾기 (ParameterSet: `FindReplace*`)
    pub fn repeat_find(&self) -> crate::error::Result<()> {
        self.h_action()?.run("RepeatFind")
    }

    /// `ReverseFind` — 거꾸로 찾기 (ParameterSet: `FindReplace*`)
    pub fn reverse_find(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ReverseFind")
    }

    /// `ReplaceDlg` — 찾아 바꾸기 대화상자를 표시합니다.
    pub fn replace_dlg(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ReplaceDlg")
    }

    /// `ReplacePrivateInfoDlg` — 개인 정보 찾아 숨기기(문자열 치환) 대화상자를 표시합니다.
    pub fn replace_private_info_dlg(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ReplacePrivateInfoDlg")
    }

    /// `DocFindInit` — 문서 찾기를 초기화합니다.
    pub fn doc_find_init(&self) -> crate::error::Result<()> {
        self.h_action()?.run("DocFindInit")
    }

    /// `DocFindNext` — 문서에서 다음 항목을 찾습니다. (ParameterSet: `FindReplace*`)
    pub fn doc_find_next(&self) -> crate::error::Result<()> {
        self.h_action()?.run("DocFindNext")
    }

    /// `DocFindEnd` — 문서 찾기를 종료합니다.
    pub fn doc_find_end(&self) -> crate::error::Result<()> {
        self.h_action()?.run("DocFindEnd")
    }

    /// `SearchPrivateInfo` — 개인 정보 찾아 감추기(암호화)를 실행합니다.
    pub fn search_private_info(&self) -> crate::error::Result<()> {
        self.h_action()?.run("SearchPrivateInfo")
    }
}
