//! SDK `HAction` — 한글 액션 실행기
//!
//! SDK 참고: HwpAutomation_2504.pdf § HAction
use crate::error::{Result, ensure};
use crate::variant::IntoVariant;

// =========================================================================
// HAction — HWP 액션 실행기
// =========================================================================

hwp_com_type!(
    /// SDK `HAction` 대응 — 한/글에서 특정 기능을 수행하기 위한 액션 오브젝트.
    ///
    /// [`HwpObject::h_action`](crate::HwpObject::h_action)으로 얻는다. SDK method 4개(`run`,
    /// `get_default`, `execute`, `popup_dialog`) 말고도, Action ID별 래퍼(`undo`, `copy` 등)가
    /// [`actions`](crate::actions) 모듈의 `impl HAction` 블록에 정의돼 있다.
    ///
    /// SDK 참고: HwpAutomation_2504.pdf § HAction
    HAction
);

impl HAction {
    /// 액션을 즉시 실행합니다.
    ///
    /// HWP COM은 실패 시 HRESULT 에러 대신 `false`를 반환하므로,
    /// 반환값이 `false`이면 `ExecutionFailed` 에러를 반환합니다.
    /// (`get_default`, `execute`, `popup_dialog`도 같습니다.)
    pub fn run(&self, action_name: &str) -> Result<()> {
        let ok: bool = self.call_with("Run", vec![action_name.into_variant()?])?;
        ensure(ok, || format!("Run(\"{action_name}\")"))
    }

    /// 액션의 기본 파라미터를 가져옵니다.
    pub fn get_default(&self, action_name: &str, pset: &crate::disp_obj::DispObj) -> Result<()> {
        let ok: bool = self.call_with(
            "GetDefault",
            vec![action_name.into_variant()?, pset.to_variant()?],
        )?;
        ensure(ok, || format!("GetDefault(\"{action_name}\")"))
    }

    /// 파라미터를 지정하여 액션을 실행합니다.
    pub fn execute(&self, action_name: &str, pset: &crate::disp_obj::DispObj) -> Result<()> {
        let ok: bool = self.call_with(
            "Execute",
            vec![action_name.into_variant()?, pset.to_variant()?],
        )?;
        ensure(ok, || format!("Execute(\"{action_name}\")"))
    }

    /// 액션의 설정 대화상자를 표시합니다.
    pub fn popup_dialog(
        &self,
        action_name: &str,
        pset: &crate::disp_obj::DispObj,
    ) -> Result<()> {
        let ok: bool = self.call_with(
            "PopupDialog",
            vec![action_name.into_variant()?, pset.to_variant()?],
        )?;
        ensure(ok, || format!("PopupDialog(\"{action_name}\")"))
    }
}
