use crate::method::h_action::HAction;

/// 앱·파일·인쇄·프레젠테이션·버전 관련 액션
///
/// SDK 참고: ActionTable_2504.pdf § App*, File*, Print*, Presentation*, Version*
impl HAction {
    // ── 앱 ──

    /// `AppQuit` — 한글을 종료합니다.
    pub fn app_quit(&self) -> crate::error::Result<()> {
        self.run("AppQuit")
    }

    /// `AppShow` — 한글 창을 화면에 표시합니다.
    pub fn app_show(&self) -> crate::error::Result<()> {
        self.run("AppShow")
    }

    /// `FrameStatusBar` — 상태 표시줄 보이기/숨기기를 토글합니다.
    pub fn frame_status_bar(&self) -> crate::error::Result<()> {
        self.run("FrameStatusBar")
    }

    /// `HwpDic` — 한글 사전 대화상자를 엽니다.
    pub fn hwp_dic(&self) -> crate::error::Result<()> {
        self.run("HwpDic")
    }

    /// `HanThDIC` — 한자·유의어 사전을 엽니다.
    pub fn han_th_dic(&self) -> crate::error::Result<()> {
        self.run("HanThDIC")
    }

    /// `SpellingCheck` — 맞춤법 검사를 실행합니다.
    pub fn spelling_check(&self) -> crate::error::Result<()> {
        self.run("SpellingCheck")
    }

    /// `ScanHFTFonts` — 한글 글꼴을 검색합니다.
    pub fn scan_hft_fonts(&self) -> crate::error::Result<()> {
        self.run("ScanHFTFonts")
    }

    /// `Preference` — 환경 설정 대화상자를 엽니다.
    pub fn preference(&self) -> crate::error::Result<()> {
        self.run("Preference")
    }

    /// `SplitMemoOpen` — 메모창을 엽니다.
    pub fn split_memo_open(&self) -> crate::error::Result<()> {
        self.run("SplitMemoOpen")
    }

    /// `ReplyMemo` — 메모 회신 (한글 2022 이상)
    pub fn reply_memo(&self) -> crate::error::Result<()> {
        self.run("ReplyMemo")
    }

    // ── 파일 ──

    /// `FileNew` — 새 빈 문서를 만듭니다.
    pub fn file_new(&self) -> crate::error::Result<()> {
        self.run("FileNew")
    }

    /// `FileOpen` — 파일 열기 대화상자를 표시합니다.
    pub fn file_open(&self) -> crate::error::Result<()> {
        self.run("FileOpen")
    }

    /// `FileSave` — 현재 문서를 저장합니다.
    pub fn file_save(&self) -> crate::error::Result<()> {
        self.run("FileSave")
    }

    /// `FileSaveAs` — 다른 이름으로 저장 대화상자를 표시합니다.
    pub fn file_save_as(&self) -> crate::error::Result<()> {
        self.run("FileSaveAs")
    }

    /// `FileClose` — 현재 문서를 닫습니다.
    pub fn file_close(&self) -> crate::error::Result<()> {
        self.run("FileClose")
    }

    /// `FileQuit` — 한글을 종료합니다. (파일 메뉴의 끝내기)
    pub fn file_quit(&self) -> crate::error::Result<()> {
        self.run("FileQuit")
    }

    /// `FilePreview` — 인쇄 미리 보기를 표시합니다.
    pub fn file_preview(&self) -> crate::error::Result<()> {
        self.run("FilePreview")
    }

    /// `FileSaveOptionDlg` — 저장 옵션 대화상자를 표시합니다.
    pub fn file_save_option_dlg(&self) -> crate::error::Result<()> {
        self.run("FileSaveOptionDlg")
    }

    /// `SaveBlockAction` — 블록 저장하기 대화상자를 표시합니다.
    pub fn save_block_action(&self) -> crate::error::Result<()> {
        self.run("SaveBlockAction")
    }

    /// `SaveHistoryItem` — 새 버전으로 저장합니다.
    pub fn save_history_item(&self) -> crate::error::Result<()> {
        self.run("SaveHistoryItem")
    }

    /// `SendBrowserText` — 브라우저로 보내기를 실행합니다.
    pub fn send_browser_text(&self) -> crate::error::Result<()> {
        self.run("SendBrowserText")
    }

    /// `SendMailAttach` — 편지 보내기 (첨부파일로) 대화상자를 표시합니다.
    pub fn send_mail_attach(&self) -> crate::error::Result<()> {
        self.run("SendMailAttach")
    }

    /// `SendMailText` — 편지 보내기 (본문으로) 대화상자를 표시합니다.
    pub fn send_mail_text(&self) -> crate::error::Result<()> {
        self.run("SendMailText")
    }

    // ── 인쇄 ──

    /// `Print` — 인쇄 대화상자를 표시합니다.
    pub fn print(&self) -> crate::error::Result<()> {
        self.run("Print")
    }

    /// `PrintSetup` — 인쇄 옵션(워터마크) 대화상자를 표시합니다.
    pub fn print_setup(&self) -> crate::error::Result<()> {
        self.run("PrintSetup")
    }

    /// `PrintToImage` — 그림으로 저장하기 대화상자를 표시합니다.
    pub fn print_to_image(&self) -> crate::error::Result<()> {
        self.run("PrintToImage")
    }

    /// `PrintToPDF` — PDF 인쇄 대화상자를 표시합니다.
    pub fn print_to_pdf(&self) -> crate::error::Result<()> {
        self.run("PrintToPDF")
    }

    // ── 프레젠테이션 ──

    /// `Presentation` — 프레젠테이션 대화상자를 표시합니다.
    pub fn presentation(&self) -> crate::error::Result<()> {
        self.run("Presentation")
    }

    /// `PresentationDelete` — 프레젠테이션을 삭제합니다.
    pub fn presentation_delete(&self) -> crate::error::Result<()> {
        self.run("PresentationDelete")
    }

    /// `PresentationRange` — 프레젠테이션 범위를 설정합니다.
    pub fn presentation_range(&self) -> crate::error::Result<()> {
        self.run("PresentationRange")
    }

    /// `PresentationSetup` — 프레젠테이션 설정 대화상자를 표시합니다.
    pub fn presentation_setup(&self) -> crate::error::Result<()> {
        self.run("PresentationSetup")
    }

    // ── 버전 ──

    /// `VersionInfo` — 버전 정보 대화상자를 표시합니다.
    pub fn version_info(&self) -> crate::error::Result<()> {
        self.run("VersionInfo")
    }

    /// `VersionSave` — 버전을 저장합니다.
    pub fn version_save(&self) -> crate::error::Result<()> {
        self.run("VersionSave")
    }

    /// `VersionDelete` — 버전 정보를 지웁니다.
    pub fn version_delete(&self) -> crate::error::Result<()> {
        self.run("VersionDelete")
    }

    /// `VersionDeleteAll` — 모든 버전 정보를 지웁니다.
    pub fn version_delete_all(&self) -> crate::error::Result<()> {
        self.run("VersionDeleteAll")
    }

    // ── 개인정보 보안 ──

    /// `PrivateInfoChangePassword` — 개인 정보 보안 암호를 변경합니다.
    pub fn private_info_change_password(&self) -> crate::error::Result<()> {
        self.run("PrivateInfoChangePassword")
    }

    /// `PrivateInfoSetPassword` — 개인 정보 보안 암호를 설정합니다.
    pub fn private_info_set_password(&self) -> crate::error::Result<()> {
        self.run("PrivateInfoSetPassword")
    }
}
