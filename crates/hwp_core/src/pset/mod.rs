//! ParameterSet 타입 — SDK method와 Action이 함께 쓰는 공용 데이터 타입
//!
//! SDK 참고: ParameterSetTable_2504.pdf, HwpAutomation_2504.pdf § HParameterSet
use crate::error::Result;

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

    /// `CharShape` 액션용 파라미터 셋
    pub fn h_char_shape(&self) -> Result<HCharShape> {
        self.get("HCharShape")
    }

    /// `ParaShape` 액션용 파라미터 셋
    pub fn h_para_shape(&self) -> Result<HParaShape> {
        self.get("HParaShape")
    }
}

// =========================================================================
// HCharShape — 글자 모양 파라미터 셋
// =========================================================================

hwp_com_type!(
    /// C++ `CHCharShape` 대응. `CharShape` 액션의 파라미터 셋입니다.
    ///
    /// SDK 참고: ParameterSetTable_2504.pdf § CharShape
    HCharShape
);

impl HCharShape {
    /// GetDefault / Execute에 전달할 파라미터 셋 핸들
    pub fn h_set(&self) -> Result<crate::disp_obj::DispObj> {
        self.get("HSet")
    }
}

// =========================================================================
// HParaShape — 문단 모양 파라미터 셋
// =========================================================================

hwp_com_type!(
    /// C++ `CHParaShape` 대응. `ParaShape` 액션의 파라미터 셋입니다.
    ///
    /// SDK 참고: ParameterSetTable_2504.pdf § ParaShape
    HParaShape
);

impl HParaShape {
    /// GetDefault / Execute에 전달할 파라미터 셋 핸들
    pub fn h_set(&self) -> Result<crate::disp_obj::DispObj> {
        self.get("HSet")
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
