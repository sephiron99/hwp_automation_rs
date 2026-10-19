//! ParameterSet 타입 — SDK method와 Action이 함께 쓰는 공용 데이터 타입
//!
//! SDK 참고: ParameterSetTable_2504.pdf, HwpAutomation_2504.pdf § HParameterSet
use crate::error::Result;

// =========================================================================
// HParameterSet — 파라미터 셋 컨테이너
// =========================================================================

hwp_com_type!(
    /// SDK `HParameterSet` 대응 — 파라메터셋 오브젝트를 관리하는 HParameterSet Object.
    ///
    /// SDK 참고: HwpAutomation_2504.pdf § IHwpObject `HParameterSet(Property)`
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

    /// `Bookmark`, `ModifyBookmark` 액션용 파라미터 셋
    pub fn h_book_mark(&self) -> Result<HBookMark> {
        self.get("HBookMark")
    }
}

// =========================================================================
// HCharShape — 글자 모양 파라미터 셋
// =========================================================================

hwp_com_type!(
    /// SDK ParameterSet `CharShape` 대응 — 글자 모양. `CharShape` 액션의 파라미터 셋입니다.
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
    /// SDK ParameterSet `ParaShape` 대응 — 문단 모양. `ParaShape` 액션의 파라미터 셋입니다.
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
    /// SDK ParameterSet `InsertText` 대응 — 텍스트 삽입. `InsertText` 액션의 파라미터 셋입니다.
    ///
    /// SDK 참고: ParameterSetTable_2504.pdf § InsertText
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
// HBookMark — 책갈피 파라미터 셋
// =========================================================================

hwp_com_type!(
    /// SDK ParameterSet `BookMark` 대응 — 책갈피. `Bookmark`, `ModifyBookmark` 액션의 파라미터 셋입니다.
    ///
    /// SDK 참고: ParameterSetTable_2504.pdf § BookMark
    HBookMark
);

impl HBookMark {
    /// GetDefault / Execute에 전달할 파라미터 셋 핸들
    pub fn h_set(&self) -> Result<crate::disp_obj::DispObj> {
        self.get("HSet")
    }

    /// 책갈피 이름
    pub fn name(&self) -> Result<String> {
        self.get("Name")
    }

    pub fn set_name(&self, name: &str) -> Result<()> {
        self.put("Name", name)
    }

    /// 책갈피 종류 (0 = 일반 책갈피, 1 = 블록 책갈피)
    pub fn type_(&self) -> Result<u32> {
        self.get("Type")
    }

    pub fn set_type(&self, type_: u32) -> Result<()> {
        self.put("Type", type_)
    }

    /// 책갈피 명령의 종류 (0 = 책갈피 생성, 1 = 책갈피로 이동, 2 = 책갈피 수정)
    pub fn command(&self) -> Result<u32> {
        self.get("Command")
    }

    pub fn set_command(&self, command: u32) -> Result<()> {
        self.put("Command", command)
    }
}
