//! `IHwpObject::MovePos` — 캐럿(caret) 위치 이동.
//!
//! SDK 참고: HwpAutomation_2504.pdf 20–21쪽.
//!
//! 원본 시그니처는
//! `BOOL MovePos([unsigned int moveID], [unsigned int para], [unsigned int pos])`
//! 로 세 개의 `unsigned int`를 받지만, `para`와 `pos`의 의미는 `moveID`에 따라
//! 달라진다. (대부분의 `moveID`는 두 값을 전혀 사용하지 않고, `moveMain`/`moveCurList`는
//! 문단 번호·문자 위치로, `moveScrPos`는 스크린 좌표로 해석된다.)
//!
//! 이 모듈은 각 `moveID`마다 의미 있는 부가 파라미터만을 담는 enum
//! [`MovePos`]를 노출한다. 그 결과 관련 없는 값을 넘기는 실수를 컴파일 타임에
//! 차단할 수 있다.

use crate::error::Result;
use crate::hwp_obj::HwpObject;
use crate::variant::IntoVariant;

/// `MovePos` 메서드에 넘길 이동 명령.
///
/// 원 SDK의 `moveID` 상수(`moveMain`, `moveCurList`, ...)에 대응되며,
/// 각 variant는 해당 `moveID`가 실제로 사용하는 부가 파라미터만을 담는다.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum MovePos {
    /// `moveMain` (0). 루트 리스트에서 `para`번째 문단의 `pos` 문자 위치로 이동.
    Main {
        /// 이동할 문단 번호.
        para: u32,
        /// 문단 내 문자 위치.
        pos: u32,
    },
    /// `moveCurList` (1). 현재 리스트에서 `para`번째 문단의 `pos` 문자 위치로 이동.
    CurList {
        /// 이동할 문단 번호.
        para: u32,
        /// 문단 내 문자 위치.
        pos: u32,
    },
    /// `moveTopOfFile` (2). 문서의 시작으로 이동.
    TopOfFile,
    /// `moveBottomOfFile` (3). 문서의 끝으로 이동.
    BottomOfFile,
    /// `moveTopOfList` (4). 현재 리스트의 시작으로 이동.
    TopOfList,
    /// `moveBottomOfList` (5). 현재 리스트의 끝으로 이동.
    BottomOfList,
    /// `moveStartOfPara` (6). 현재 위치한 문단의 시작으로 이동.
    StartOfPara,
    /// `moveEndOfPara` (7). 현재 위치한 문단의 끝으로 이동.
    EndOfPara,
    /// `moveStartOfWord` (8). 현재 위치한 단어의 시작으로 이동. (현재 리스트만 대상)
    StartOfWord,
    /// `moveEndOfWord` (9). 현재 위치한 단어의 끝으로 이동. (현재 리스트만 대상)
    EndOfWord,
    /// `moveNextPara` (10). 다음 문단의 시작으로 이동. (현재 리스트만 대상)
    NextPara,
    /// `movePrevPara` (11). 앞 문단의 끝으로 이동. (현재 리스트만 대상)
    PrevPara,
    /// `moveNextPos` (12). 한 글자 앞으로 이동. (서브 리스트를 넘나들 수 있음)
    NextPos,
    /// `movePrevPos` (13). 한 글자 뒤로 이동. (서브 리스트를 넘나들 수 있음)
    PrevPos,
    /// `moveNextPosEx` (14). 한 글자 앞으로 이동.
    /// (서브 리스트를 넘나들 수 있으며 머리말/꼬리말, 각주/미주, 글상자 포함)
    NextPosEx,
    /// `movePrevPosEx` (15). 한 글자 뒤로 이동.
    /// (서브 리스트를 넘나들 수 있으며 머리말/꼬리말, 각주/미주, 글상자 포함)
    PrevPosEx,
    /// `moveNextChar` (16). 한 글자 앞으로 이동. (현재 리스트만 대상)
    NextChar,
    /// `movePrevChar` (17). 한 글자 뒤로 이동. (현재 리스트만 대상)
    PrevChar,
    /// `moveNextWord` (18). 한 단어 앞으로 이동. (현재 리스트만 대상)
    NextWord,
    /// `movePrevWord` (19). 한 단어 뒤로 이동. (현재 리스트만 대상)
    PrevWord,
    /// `moveNextLine` (20). 한 줄 위로 이동.
    NextLine,
    /// `movePrevLine` (21). 한 줄 아래로 이동.
    PrevLine,
    /// `moveStartOfLine` (22). 현재 위치한 줄의 시작으로 이동.
    StartOfLine,
    /// `moveEndOfLine` (23). 현재 위치한 줄의 끝으로 이동.
    EndOfLine,
    /// `moveParentList` (24). 한 레벨 상위로 이동.
    ParentList,
    /// `moveTopLevelList` (25). 탑레벨 리스트로 이동.
    TopLevelList,
    /// `moveRootList` (26). 루트 리스트로 이동.
    ///
    /// 추가 설명: 현재 루트 리스트에 위치해 있어 더 이상 상위 리스트가 없을 때는
    /// 위치 이동 없이 리턴한다. 이동 후의 위치는 상위 리스트에서 서브리스트가
    /// 속한 컨트롤 코드가 위치한 곳이다. 위치 이동 시 셀렉션은 무조건 풀린다.
    RootList,
    /// `moveLeftOfCell` (100). 현재 캐럿이 위치한 셀의 왼쪽.
    LeftOfCell,
    /// `moveRightOfCell` (101). 현재 캐럿이 위치한 셀의 오른쪽.
    RightOfCell,
    /// `moveUpOfCell` (102). 현재 캐럿이 위치한 셀의 위쪽.
    UpOfCell,
    /// `moveDownOfCell` (103). 현재 캐럿이 위치한 셀의 아래쪽.
    DownOfCell,
    /// `moveStartOfCell` (104). 현재 캐럿이 위치한 셀에서 행(row)의 시작.
    StartOfCell,
    /// `moveEndOfCell` (105). 현재 캐럿이 위치한 셀에서 행(row)의 끝.
    EndOfCell,
    /// `moveTopOfCell` (106). 현재 캐럿이 위치한 셀에서 열(column)의 시작.
    TopOfCell,
    /// `moveBottomOfCell` (107). 현재 캐럿이 위치한 셀에서 열(column)의 끝.
    BottomOfCell,
    /// `moveScrPos` (200). 한/글 문서창에서의 screen 좌표로 캐럿 위치를 설정한다.
    ///
    /// `(x, y)`는 마우스 커서의 좌표와 동일하게 넘기면 된다.
    /// (원 SDK에서는 `LOWORD = x`, `HIWORD = y`로 `para`에 packing해서 넘긴다.)
    ScrPos {
        /// 스크린 x 좌표.
        x: u16,
        /// 스크린 y 좌표.
        y: u16,
    },
    /// `moveScanPos` (201). `GetText()` 실행 후 위치로 이동.
    ///
    /// 문서를 검색하는 중에 캐럿을 이동시키려 할 경우에만 사용 가능하다.
    ScanPos,
}

impl MovePos {
    /// 이 이동 명령에 대응되는 `moveID` 수치값을 반환한다.
    pub fn id(self) -> u32 {
        match self {
            MovePos::Main { .. } => 0,
            MovePos::CurList { .. } => 1,
            MovePos::TopOfFile => 2,
            MovePos::BottomOfFile => 3,
            MovePos::TopOfList => 4,
            MovePos::BottomOfList => 5,
            MovePos::StartOfPara => 6,
            MovePos::EndOfPara => 7,
            MovePos::StartOfWord => 8,
            MovePos::EndOfWord => 9,
            MovePos::NextPara => 10,
            MovePos::PrevPara => 11,
            MovePos::NextPos => 12,
            MovePos::PrevPos => 13,
            MovePos::NextPosEx => 14,
            MovePos::PrevPosEx => 15,
            MovePos::NextChar => 16,
            MovePos::PrevChar => 17,
            MovePos::NextWord => 18,
            MovePos::PrevWord => 19,
            MovePos::NextLine => 20,
            MovePos::PrevLine => 21,
            MovePos::StartOfLine => 22,
            MovePos::EndOfLine => 23,
            MovePos::ParentList => 24,
            MovePos::TopLevelList => 25,
            MovePos::RootList => 26,
            MovePos::LeftOfCell => 100,
            MovePos::RightOfCell => 101,
            MovePos::UpOfCell => 102,
            MovePos::DownOfCell => 103,
            MovePos::StartOfCell => 104,
            MovePos::EndOfCell => 105,
            MovePos::TopOfCell => 106,
            MovePos::BottomOfCell => 107,
            MovePos::ScrPos { .. } => 200,
            MovePos::ScanPos => 201,
        }
    }

    /// SDK 원형 시그니처의 `(para, pos)` raw 인자쌍을 반환한다.
    ///
    /// - `Main`/`CurList`: 그대로 문단 번호와 문자 위치.
    /// - `ScrPos`: `para`에 `LOWORD = x`, `HIWORD = y`로 packing. `pos`는 0.
    /// - 그 외: `(0, 0)`. SDK가 무시한다.
    fn para_pos(self) -> (u32, u32) {
        match self {
            MovePos::Main { para, pos } | MovePos::CurList { para, pos } => (para, pos),
            MovePos::ScrPos { x, y } => ((x as u32) | ((y as u32) << 16), 0),
            _ => (0, 0),
        }
    }
}

impl HwpObject {
    /// 캐럿(caret)의 위치를 이동한다.
    ///
    /// SDK 참고: HwpAutomation_2504.pdf 20–21쪽 `MovePos(moveID, para, pos)`.
    ///
    /// `moveID`에 따라 필요한 부가 파라미터가 [`MovePos`] enum의 각 variant에
    /// 이미 담겨 있으므로, 호출자는 한 개의 값만 넘기면 된다.
    ///
    /// # 반환값
    /// SDK의 `MovePos`는 `BOOL`을 돌려주는데, 이동에 실패하면(`false`) HRESULT
    /// 에러 대신 조용히 `false`만 리턴한다. 이 래퍼는 `false`를 명시적으로
    /// [`HwpError::ExecutionFailed`](crate::error::HwpError::ExecutionFailed)로
    /// 변환하여, `?`로 바로 실패를 감지할 수 있게 한다.
    ///
    /// # 예시
    /// ```ignore
    /// hwp.move_pos(MovePos::TopOfFile)?;
    /// hwp.move_pos(MovePos::CurList { para: 3, pos: 0 })?;
    /// hwp.move_pos(MovePos::ScrPos { x: 120, y: 240 })?;
    /// ```
    pub fn move_pos(&self, to: MovePos) -> Result<()> {
        let (para, pos) = to.para_pos();
        let ok: bool = self.call_with(
            "MovePos",
            vec![
                to.id().into_variant()?,
                para.into_variant()?,
                pos.into_variant()?,
            ],
        )?;
        if !ok {
            return Err(crate::error::HwpError::ExecutionFailed(format!(
                "MovePos({to:?})"
            )));
        }
        Ok(())
    }
}
