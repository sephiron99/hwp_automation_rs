/// 편집 용지·바탕쪽·구역·줄 번호 관련 액션
///
/// SDK 참고: ActionTable_2504.pdf § Page*, MP*, SetLineNumbers, OutlineNumber 등
use crate::hwp_obj::HwpObject;

impl HwpObject {
    // ── 편집 용지 ──

    /// `PageSetup` — 편집 용지 대화상자 (ParameterSet: `SecDef`)
    pub fn page_setup(&self) -> crate::error::Result<()> {
        self.h_action()?.run("PageSetup")
    }

    /// `PageSetupDL` — 편집 용지 (쪽 여백 설정, ParameterSet: `SecDef`)
    pub fn page_setup_dl(&self) -> crate::error::Result<()> {
        self.h_action()?.run("PageSetupDL")
    }

    /// `PageMarginSetup` — 편집 용지 (쪽 여백 설정, 한글 2022 이상, ParameterSet: `SecDef`)
    pub fn page_margin_setup(&self) -> crate::error::Result<()> {
        self.h_action()?.run("PageMarginSetup")
    }

    /// `PageLandscape` — 용지 넓게 (가로 방향, ParameterSet: `SecDef`)
    pub fn page_landscape(&self) -> crate::error::Result<()> {
        self.h_action()?.run("PageLandscape")
    }

    /// `PagePortrait` — 용지 좁게 (세로 방향, ParameterSet: `SecDef`)
    pub fn page_portrait(&self) -> crate::error::Result<()> {
        self.h_action()?.run("PagePortrait")
    }

    // ── 쪽 테두리·배경 ──

    /// `PageBorder` — 쪽 테두리/배경 대화상자 (ParameterSet: `SecDef`)
    pub fn page_border(&self) -> crate::error::Result<()> {
        self.h_action()?.run("PageBorder")
    }

    /// `PageBorderTab` — 쪽 테두리/배경 (항상 테두리 탭이 선택되어 보임, ParameterSet: `SecDef`)
    pub fn page_border_tab(&self) -> crate::error::Result<()> {
        self.h_action()?.run("PageBorderTab")
    }

    /// `PageFillTab` — 쪽 테두리/배경 (항상 채우기 탭이 선택되어 보임, ParameterSet: `SecDef`)
    pub fn page_fill_tab(&self) -> crate::error::Result<()> {
        self.h_action()?.run("PageFillTab")
    }

    // ── 쪽 번호 ──

    /// `PageNumPos` — 쪽 번호 매기기 (ParameterSet: `PageNumPos`)
    pub fn page_num_pos(&self) -> crate::error::Result<()> {
        self.h_action()?.run("PageNumPos")
    }

    /// `PageNumPosModify` — 쪽 번호 매기기 고치기 (ParameterSet: `PageNumPos`)
    pub fn page_num_pos_modify(&self) -> crate::error::Result<()> {
        self.h_action()?.run("PageNumPosModify")
    }

    // ── 쪽 숨기기 ──

    /// `PageHiding` — 감추기 (ParameterSet: `PageHiding`)
    pub fn page_hiding(&self) -> crate::error::Result<()> {
        self.h_action()?.run("PageHiding")
    }

    /// `PageHidingModify` — 감추기 고치기 (ParameterSet: `PageHiding`)
    pub fn page_hiding_modify(&self) -> crate::error::Result<()> {
        self.h_action()?.run("PageHidingModify")
    }

    // ── 줄 번호 ──

    /// `SetLineNumbers` — 줄 번호 넣기 (ParameterSet: `SecDef`)
    pub fn set_line_numbers(&self) -> crate::error::Result<()> {
        self.h_action()?.run("SetLineNumbers")
    }

    /// `ShowLineNumbers` — 줄 번호 넣기 (ParameterSet: `SecDef`)
    pub fn show_line_numbers(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ShowLineNumbers")
    }

    /// `SuppressLineNumbers` — 줄 번호 넣기 (현재 문단 숨김)
    pub fn suppress_line_numbers(&self) -> crate::error::Result<()> {
        self.h_action()?.run("SuppressLineNumbers")
    }

    // ── 개요 번호 ──

    /// `OutlineNumber` — 개요번호 (ParameterSet: `SecDef`)
    pub fn outline_number(&self) -> crate::error::Result<()> {
        self.h_action()?.run("OutlineNumber")
    }

    // ── 다단 ──

    /// `MultiColumn` — 다단 설정 (ParameterSet: `ColDef`)
    pub fn multi_column(&self) -> crate::error::Result<()> {
        self.h_action()?.run("MultiColumn")
    }

    // ── 바탕쪽 (MasterPage) ──

    /// `MPBreakNewSection` — 새 구역 만들기 (바탕쪽 편집 상태에서, ParameterSet: `MasterPage`)
    pub fn mp_break_new_section(&self) -> crate::error::Result<()> {
        self.h_action()?.run("MPBreakNewSection")
    }

    /// `MPCopyFromOtherSection` — 바탕쪽 가져오기-다른 구역의 바탕쪽 종류와 내용을 복사 (ParameterSet: `Masterpage`)
    pub fn mp_copy_from_other_section(&self) -> crate::error::Result<()> {
        self.h_action()?.run("MPCopyFromOtherSection")
    }

    /// `MPSectionToNext` — 이후 구역으로 이동
    pub fn mp_section_to_next(&self) -> crate::error::Result<()> {
        self.h_action()?.run("MPSectionToNext")
    }

    /// `MPSectionToPrevious` — 이전 구역으로 이동
    pub fn mp_section_to_previous(&self) -> crate::error::Result<()> {
        self.h_action()?.run("MPSectionToPrevious")
    }

    /// `MPShowMarginBorder` — 여백 보기 (바탕쪽 편집 상태에서)
    pub fn mp_show_margin_border(&self) -> crate::error::Result<()> {
        self.h_action()?.run("MPShowMarginBorder")
    }
}
