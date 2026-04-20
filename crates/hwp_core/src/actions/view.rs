/// 보기 옵션·화면 배율 관련 액션
///
/// SDK 참고: ActionTable_2504.pdf § View*, ViewZoom*, ViewOption*
use crate::hwp_obj::HwpObject;

impl HwpObject {
    // ── 화면 배율 ──

    /// `ViewZoom` — 화면 확대 대화상자 (Ribbon, ParameterSet: `ViewProperties`)
    pub fn view_zoom(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ViewZoom")
    }

    /// `ViewZoomNormal` — 화면 확대: 정상 (ParameterSet: `ViewProperties`)
    pub fn view_zoom_normal(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ViewZoomNormal")
    }

    /// `ViewZoomFitPage` — 화면 확대: 페이지에 맞춤 (ParameterSet: `ViewProperties`)
    pub fn view_zoom_fit_page(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ViewZoomFitPage")
    }

    /// `ViewZoomFitWidth` — 화면 확대: 폭에 맞춤 (ParameterSet: `ViewProperties`)
    pub fn view_zoom_fit_width(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ViewZoomFitWidth")
    }

    /// `ViewZoomLock` — 화면 잠금
    pub fn view_zoom_lock(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ViewZoomLock")
    }

    // ── 격자 ──

    /// `ViewGridOption` — 격자 설정 (ParameterSet: `GridInfo`)
    pub fn view_grid_option(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ViewGridOption")
    }

    /// `ViewShowGrid` — 격자 보이기 (ParameterSet: `GridInfo`)
    pub fn view_show_grid(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ViewShowGrid")
    }

    // ── 보기 옵션 ──

    /// `ViewIdiom` — 상용구 보기
    pub fn view_idiom(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ViewIdiom")
    }

    /// `ViewOptionCtrlMark` — 조판 부호 보이기/숨기기
    pub fn view_option_ctrl_mark(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ViewOptionCtrlMark")
    }

    /// `ViewOptionParaMark` — 문단 부호 보이기/숨기기
    pub fn view_option_para_mark(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ViewOptionParaMark")
    }

    /// `ViewOptionGuideLine` — 안내선 보이기/숨기기
    pub fn view_option_guide_line(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ViewOptionGuideLine")
    }

    /// `ViewOptionPaper` — 쪽 윤곽 보이기/숨기기
    pub fn view_option_paper(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ViewOptionPaper")
    }

    /// `ViewOptionPicture` — 그림 보이기/숨기기 (보기-그림 메뉴와 동일)
    pub fn view_option_picture(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ViewOptionPicture")
    }

    /// `ViewOptionRevision` — 교정부호 보이기/숨기기 (보기-교정부호 메뉴와 동일)
    pub fn view_option_revision(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ViewOptionRevision")
    }

    /// `ViewOptionMemo` — 메모 보이기/숨기기 (보기-메모-메모 보이기/숨기기 메뉴와 동일)
    pub fn view_option_memo(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ViewOptionMemo")
    }

    /// `ViewOptionMemoGuideline` — 메모 안내선 표시
    pub fn view_option_memo_guideline(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ViewOptionMemoGuideline")
    }

    /// `ViewOptionColor` — 컬러로 보기 (회색조 보기 되돌리기 액션)
    pub fn view_option_color(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ViewOptionColor")
    }

    /// `ViewOptionColorCustom` — 사용자색 보기
    pub fn view_option_color_custom(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ViewOptionColorCustom")
    }

    /// `ViewOptionColorCustomOption` — 사용자색 설정
    pub fn view_option_color_custom_option(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ViewOptionColorCustomOption")
    }

    /// `ViewOptionGray` — 회색조 보기
    pub fn view_option_gray(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ViewOptionGray")
    }

    /// `ViewOptionPronounce` — 한자/일어 발음 표시 Toggle (ParameterSet: `PronounceInfo`)
    pub fn view_option_pronounce(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ViewOptionPronounce")
    }

    /// `ViewOptionPronounceSetting` — 한자/일어 발음 표시 설정 (ParameterSet: `PronounceInfo`)
    pub fn view_option_pronounce_setting(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ViewOptionPronounceSetting")
    }

    // ── 변경 추적 보기 ──

    /// `ViewOptionTrackChange` — 변경추적 보기
    pub fn view_option_track_change(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ViewOptionTrackChange")
    }

    /// `ViewOptionTrackChangeFinal` — 변경추적 보기: 최종본 보기
    pub fn view_option_track_change_final(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ViewOptionTrackChangeFinal")
    }

    /// `ViewOptionTrackChangeFinalMemo` — 변경추적 보기: 메모 및 변경 내용 최종본
    pub fn view_option_track_change_final_memo(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ViewOptionTrackChangeFinalMemo")
    }

    /// `ViewOptionTrackChangeInline` — 변경추적 보기: 안내문에 표시
    pub fn view_option_track_change_inline(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ViewOptionTrackChangeInline")
    }

    /// `ViewOptionTrackChangeInsertDelete` — 변경추적 보기: 삽입 및 삭제
    pub fn view_option_track_change_insert_delete(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ViewOptionTrackChangeInsertDelete")
    }

    /// `ViewOptionTrackChangeOriginal` — 변경추적 보기: 원본 보기
    pub fn view_option_track_change_original(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ViewOptionTrackChangeOriginal")
    }

    /// `ViewOptionTrackChangeOriginalMemo` — 변경추적 보기: 메모 및 변경 내용 원본
    pub fn view_option_track_change_original_memo(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ViewOptionTrackChangeOriginalMemo")
    }

    /// `ViewOptionTrackChangeShape` — 변경추적 보기: 서식
    pub fn view_option_track_change_shape(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ViewOptionTrackChangeShape")
    }

    /// `ViewOptionTrackChnageInfo` — 변경추적 보기: 변경 내용 보기 (PDF 오타: Chnage)
    pub fn view_option_track_change_info(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ViewOptionTrackChnageInfo")
    }
}
