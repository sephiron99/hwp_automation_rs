/// 변경 추적 관련 액션
///
/// SDK 참고: ActionTable_2504.pdf § TrackChange*
use crate::hwp_obj::HwpObject;

impl HwpObject {
    /// `TrackChangeOption` — 변경 내용 추적 설정 (ParameterSet: `TrackChange`)
    pub fn track_change_option(&self) -> crate::error::Result<()> {
        self.h_action()?.run("TrackChangeOption")
    }

    /// `TrackChangeProtection` — 변경추적 보호 (ParameterSet: `Password`)
    pub fn track_change_protection(&self) -> crate::error::Result<()> {
        self.h_action()?.run("TrackChangeProtection")
    }

    /// `TrackChangeAuthor` — 변경추적: 사용자 이름 변경
    pub fn track_change_author(&self) -> crate::error::Result<()> {
        self.h_action()?.run("TrackChangeAuthor")
    }

    // ── 변경 내용 적용 ──

    /// `TrackChangeApply` — 변경추적: 변경내용 적용
    pub fn track_change_apply(&self) -> crate::error::Result<()> {
        self.h_action()?.run("TrackChangeApply")
    }

    /// `TrackChangeApplyAll` — 변경추적: 문서에서 변경내용 모두 적용
    pub fn track_change_apply_all(&self) -> crate::error::Result<()> {
        self.h_action()?.run("TrackChangeApplyAll")
    }

    /// `TrackChangeApplyViewAll` — 변경추적: 표시된 변경내용 모두 적용
    pub fn track_change_apply_view_all(&self) -> crate::error::Result<()> {
        self.h_action()?.run("TrackChangeApplyViewAll")
    }

    /// `TrackChangeApplyNext` — 변경추적: 적용 후 다음으로 이동
    pub fn track_change_apply_next(&self) -> crate::error::Result<()> {
        self.h_action()?.run("TrackChangeApplyNext")
    }

    /// `TrackChangeApplyPrev` — 변경추적: 적용 후 이전으로 이동
    pub fn track_change_apply_prev(&self) -> crate::error::Result<()> {
        self.h_action()?.run("TrackChangeApplyPrev")
    }

    // ── 변경 내용 취소 ──

    /// `TrackChangeCancel` — 변경추적: 변경내용 취소
    pub fn track_change_cancel(&self) -> crate::error::Result<()> {
        self.h_action()?.run("TrackChangeCancel")
    }

    /// `TrackChangeCancelAll` — 변경추적: 문서에서 변경내용 모두 취소
    pub fn track_change_cancel_all(&self) -> crate::error::Result<()> {
        self.h_action()?.run("TrackChangeCancelAll")
    }

    /// `TrackChangeCancelViewAll` — 변경추적: 표시된 변경내용 모두 취소
    pub fn track_change_cancel_view_all(&self) -> crate::error::Result<()> {
        self.h_action()?.run("TrackChangeCancelViewAll")
    }

    /// `TrackChangeCancelNext` — 변경추적: 취소 후 다음으로 이동
    pub fn track_change_cancel_next(&self) -> crate::error::Result<()> {
        self.h_action()?.run("TrackChangeCancelNext")
    }

    /// `TrackChangeCancelPrev` — 변경추적: 취소 후 이전으로 이동
    pub fn track_change_cancel_prev(&self) -> crate::error::Result<()> {
        self.h_action()?.run("TrackChangeCancelPrev")
    }

    // ── 이동 ──

    /// `TrackChangeNext` — 변경추적: 다음 변경내용
    pub fn track_change_next(&self) -> crate::error::Result<()> {
        self.h_action()?.run("TrackChangeNext")
    }

    /// `TrackChangePrev` — 변경추적: 이전 변경내용
    pub fn track_change_prev(&self) -> crate::error::Result<()> {
        self.h_action()?.run("TrackChangePrev")
    }
}
