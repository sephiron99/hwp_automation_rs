//! 문서 안의 위치 — 리스트·문단·글자
//!
//! SDK는 문서 안의 위치를 세 값으로 나타낸다.
//!
//! - `list`: 리스트 아이디. 본문은 0이고, 표 칸·머리말·각주 등은 각자 다른 번호다. 번호를
//!   매기는 규칙은 SDK에 없다.
//! - `para`: 그 리스트 안의 문단 번호
//! - `pos`: 문단 안의 글자 단위 위치. 사람이 보는 글자 수가 아니라 한글 내부 단위라,
//!   표·그림 같은 컨트롤 하나가 8칸을 차지한다.
//!
//! `GetPos`·`SetPos`·`GetSelectedPos`·ParameterSet `ListParaPos`는 세 값을 모두 쓰고,
//! `SelectText`·`MovePos`·`InitScan`은 리스트 없이 "현재 리스트 안"의 문단·글자만 받는다.
//! 그래서 두 경우를 [`ListParaPos`]와 [`ParaPos`]로 나눠, 다른 리스트의 위치를 현재
//! 리스트 API에 넘기는 실수가 [`ListParaPos::para_pos`]로 드러나게 한다.
//!
//! 화면 배치 기준인 `KeyIndicator`의 쪽·줄·칸은 다른 개념이라 여기서 다루지 않는다.
//!
//! SDK 참고: HwpAutomation_2504.pdf § IHwpObject `GetPos`, `GetSelectedPos`, `SelectText`

/// 문서 안의 한 위치 (SDK: `list`, `para`, `pos`, ParameterSet `ListParaPos`)
///
/// 서로 다른 리스트의 위치는 앞뒤를 비교할 수 없고 리스트 번호도 문서 순서가 아니어서
/// 순서 비교(`Ord`)는 두지 않는다.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ListParaPos {
    /// 리스트 아이디. 본문은 0
    pub list: i32,
    /// 리스트 안의 문단 번호
    pub para: i32,
    /// 문단 안의 글자 단위 위치(한글 내부 단위)
    pub pos: i32,
}

impl ListParaPos {
    pub const fn new(list: i32, para: i32, pos: i32) -> Self {
        Self { list, para, pos }
    }

    /// 리스트를 떼고 문단·글자만 남긴다. 리스트 없이 받는 API(`SelectText` 등)에 넘길 때
    /// 쓴다. 그 API는 현재 리스트를 기준으로 하므로, 이 위치의 리스트가 현재 리스트여야 한다.
    pub const fn para_pos(self) -> ParaPos {
        ParaPos {
            para: self.para,
            pos: self.pos,
        }
    }
}

/// 같은 리스트 안의 위치. 리스트 없이 받는 `SelectText`·`MovePos`·`InitScan`용
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ParaPos {
    /// 문단 번호
    pub para: i32,
    /// 문단 안의 글자 단위 위치(한글 내부 단위)
    pub pos: i32,
}

impl ParaPos {
    pub const fn new(para: i32, pos: i32) -> Self {
        Self { para, pos }
    }
}

/// 블록 범위 (SDK: `GetSelectedPos`의 `slist`..`epos`)
///
/// `start`가 문서상 앞쪽, `end`가 뒤쪽이며 선택한 방향과 관계없다(한글 2018에서 확인).
/// `end` 자리의 글자는 범위에 들어가지 않는다.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ListParaRange {
    pub start: ListParaPos,
    pub end: ListParaPos,
}
