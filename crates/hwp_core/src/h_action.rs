/// SDK `HAction` — 한글 액션 실행기와 관련 파라미터 타입
///
/// SDK 참고: HwpAutomation_2504.pdf § HAction, HParameterSet, HInsertText
use crate::error::Result;
use crate::hwp_obj::HwpObject;
use crate::variant::IntoVariant;

// =========================================================================
// HAction — HWP 액션 실행기
// =========================================================================

hwp_com_type!(
    /// C++ `CHAction` 대응. HWP 액션을 실행합니다.
    HAction
);

impl HAction {
    /// 액션을 즉시 실행합니다.
    ///
    /// HWP COM은 실패 시 HRESULT 에러 대신 `false`를 반환하므로,
    /// 반환값이 `false`이면 `ExecutionFailed` 에러를 반환합니다.
    pub fn run(&self, action_name: &str) -> Result<()> {
        let ok: bool = self.call_with("Run", vec![action_name.into_variant()?])?;
        if !ok {
            return Err(crate::error::HwpError::ExecutionFailed(format!(
                "Run(\"{action_name}\")"
            )));
        }
        Ok(())
    }

    /// 액션의 기본 파라미터를 가져옵니다.
    pub fn get_default(&self, action_name: &str, pset: &crate::disp_obj::DispObj) -> Result<()> {
        let ok: bool = self.call_with(
            "GetDefault",
            vec![action_name.into_variant()?, pset.to_variant()?],
        )?;
        if !ok {
            return Err(crate::error::HwpError::ExecutionFailed(format!(
                "GetDefault(\"{action_name}\")"
            )));
        }
        Ok(())
    }

    /// 파라미터를 지정하여 액션을 실행합니다.
    pub fn execute(&self, action_name: &str, pset: &crate::disp_obj::DispObj) -> Result<()> {
        let ok: bool = self.call_with(
            "Execute",
            vec![action_name.into_variant()?, pset.to_variant()?],
        )?;
        if !ok {
            return Err(crate::error::HwpError::ExecutionFailed(format!(
                "Execute(\"{action_name}\")"
            )));
        }
        Ok(())
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
        if !ok {
            return Err(crate::error::HwpError::ExecutionFailed(format!(
                "PopupDialog(\"{action_name}\")"
            )));
        }
        Ok(())
    }
}

// =========================================================================
// HParameterSet — 파라미터 셋 컨테이너
// =========================================================================

hwp_com_type!(
    /// C++ `CHParameterSet` 대응. 각종 액션 파라미터 셋의 컨테이너입니다.
    HParameterSet
);

impl HParameterSet {
    pub fn h_insert_text(&self) -> Result<HInsertText> {
        self.get("HInsertText")
    }
}

// =========================================================================
// HInsertText — 텍스트 삽입 파라미터
// =========================================================================

hwp_com_type!(
    /// C++ `CHInsertText` 대응. 텍스트 삽입 액션의 파라미터입니다.
    HInsertText
);

impl HInsertText {
    /// 파라미터 셋 객체 (HAction.GetDefault / Execute에 전달)
    pub fn h_set(&self) -> Result<crate::disp_obj::DispObj> {
        self.get("HSet")
    }

    pub fn text(&self) -> Result<String> {
        self.get("Text")
    }

    pub fn set_text(&self, text: &str) -> Result<()> {
        self.put("Text", text)
    }
}

// =========================================================================
// IHwpObject — HAction 프로퍼티 관련 메서드
// =========================================================================

impl HwpObject {
    /// `HAction` 프로퍼티 — 액션 실행기
    pub fn h_action(&self) -> crate::error::Result<HAction> {
        self.get("HAction")
    }

    /// `HParameterSet` 프로퍼티 — 파라미터 셋 컨테이너
    pub fn h_parameter_set(&self) -> crate::error::Result<HParameterSet> {
        self.get("HParameterSet")
    }

    /// `Select` — 현재 커서 위치의 단어를 선택합니다.
    ///
    /// SDK 액션 ID: `"Select"` (파라미터 없음)
    pub fn select(&self) -> crate::error::Result<()> {
        let action = self.h_action()?;
        action.run("Select")?;
        Ok(())
    }

    /// `MoveWordBegin` — 커서를 현재 단어의 시작으로 이동합니다.
    ///
    /// SDK 액션 ID: `"MoveWordBegin"` (파라미터 없음)
    pub fn move_word_begin(&self) -> crate::error::Result<()> {
        let action = self.h_action()?;
        action.run("MoveWordBegin")?;
        Ok(())
    }

    /// `InsertText` — 현재 커서 위치에 텍스트를 삽입합니다.
    ///
    /// SDK 액션 ID: `"InsertText"`, 파라미터: `HInsertText.Text`
    pub fn insert_text(&self, text: &str) -> crate::error::Result<()> {
        let action = self.h_action()?;
        let pset = self.h_parameter_set()?;
        let insert_text = pset.h_insert_text()?;
        let hset = insert_text.h_set()?;
        action.get_default("InsertText", &hset)?;
        insert_text.set_text(text)?;
        action.execute("InsertText", &hset)?;
        Ok(())
    }
}
