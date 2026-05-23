/// 현재 커서 위치의 문단 모양 정보 (FormatSpec 캡처용)
///
/// SDK 참고: ParameterSetTable_2504.pdf § ParaShape (91번)
/// `HAction.GetDefault("ParaShape", hset)` + `HParameterSet.HParaShape` 로 조회합니다.
#[derive(Debug, Clone, Default)]
pub struct ParaShape {
    /// 정렬 방식 (0=양쪽, 1=왼쪽, 2=오른쪽, 3=가운데, 4=배분, 5=나눔)
    pub align_type: u32,
    /// 줄 간격 종류 (0=글자에 따라, 1=고정값, 2=여백만 지정)
    pub line_spacing_type: u32,
    /// 줄 간격 값 (LineSpacingType=0이면 %, 1·2이면 HWPUNIT)
    pub line_spacing: i32,
    /// 왼쪽 여백 (URC)
    pub left_margin: i32,
    /// 오른쪽 여백 (URC)
    pub right_margin: i32,
    /// 들여쓰기/내어쓰기 (URC, 양수=들여쓰기, 음수=내어쓰기)
    pub indentation: i32,
    /// 문단 위 간격 (URC)
    pub prev_spacing: i32,
    /// 문단 아래 간격 (URC)
    pub next_spacing: i32,
}

// =========================================================================
// IHwpObject — ParaShape 조회 메서드
// =========================================================================

impl crate::hwp_obj::HwpObject {
    /// 현재 커서 위치의 문단 모양을 읽어 [`ParaShape`]로 반환합니다.
    ///
    /// 내부적으로 `HAction.GetDefault("ParaShape", hset)`을 호출하여
    /// `HParameterSet.HParaShape`에서 각 필드를 읽습니다.
    ///
    /// # 용도 (eat_hwp_raw)
    /// `FormatSpec` 캡처 — 삽입 후 원본 문단 모양을 재적용하거나
    /// 열 폭(column width) 추정에 활용합니다.
    pub fn current_para_shape(&self) -> crate::error::Result<ParaShape> {
        let action = self.h_action()?;
        let pset = self.h_parameter_set()?;
        let ps = pset.h_para_shape()?;
        let hset = ps.h_set()?;
        action.get_default("ParaShape", &hset)?;

        Ok(ParaShape {
            align_type: ps.get("AlignType").unwrap_or(0),
            line_spacing_type: ps.get("LineSpacingType").unwrap_or(0),
            line_spacing: ps.get("LineSpacing").unwrap_or(160),
            left_margin: ps.get("LeftMargin").unwrap_or(0),
            right_margin: ps.get("RightMargin").unwrap_or(0),
            indentation: ps.get("Indentation").unwrap_or(0),
            prev_spacing: ps.get("PrevSpacing").unwrap_or(0),
            next_spacing: ps.get("NextSpacing").unwrap_or(0),
        })
    }
}
