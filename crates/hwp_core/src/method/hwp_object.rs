/// SDK `IHwpObject` — 한글 OLE Automation 최상위 객체의 메서드 구현
///
/// SDK 참고: HwpAutomation_2504.pdf § IHwpObject
use crate::com_util;
use crate::error::ensure;
use crate::hwp_obj::HwpObject;
use crate::variant::{FromVariant, IntoVariant};

// =========================================================================
// GetTextFile 관련 타입
// =========================================================================

/// `GetTextFile`의 `format` 파라미터
///
/// SDK 문자열: "HWP", "HWPML2X", "HTML", "UNICODE", "TEXT"
pub enum GetTextFileFormat {
    /// 유니코드 텍스트
    Unicode,
    /// 일반 텍스트 (한자·특수문자 손실 가능)
    Text,
    /// HTML
    Html,
    /// HWP 네이티브 포맷 (BASE64 인코딩)
    Hwp,
    /// HWPML2X (문서 정보 완전 보존)
    HwpMl2X,
}

impl GetTextFileFormat {
    fn as_str(&self) -> &'static str {
        match self {
            GetTextFileFormat::Unicode => "UNICODE",
            GetTextFileFormat::Text => "TEXT",
            GetTextFileFormat::Html => "HTML",
            GetTextFileFormat::Hwp => "HWP",
            GetTextFileFormat::HwpMl2X => "HWPML2X",
        }
    }
}

// =========================================================================
// InitScan 관련 상수
// =========================================================================

/// `InitScan` `option` 파라미터 (SDK: `maskNormal`, `maskChar`, `maskInline`, `maskCtrl`)
pub mod mask {
    /// 본문 텍스트만 (서브 리스트 제외)
    pub const NORMAL: u32 = 0x00;
    /// Char 타입 컨트롤 포함 (강제 줄바꿈, 하이픈 등)
    pub const CHAR: u32 = 0x01;
    /// Inline 타입 컨트롤 포함
    pub const INLINE: u32 = 0x02;
    /// Extended 타입 컨트롤 포함 (머리말, 각주 등)
    pub const CTRL: u32 = 0x04;
}

/// ScanStartingPosition. `InitScan` `range` 시작 위치 (SDK: `scanSposXxx`)
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u32)]
pub enum ScanSpos {
    /// 캐럿 위치부터 (0x0000)
    Current = 0x0000,
    /// 특정 위치부터 (0x0010)
    Specified = 0x0010,
    /// 줄의 시작부터 (0x0020)
    Line = 0x0020,
    /// 문단의 시작부터 (0x0030)
    Paragraph = 0x0030,
    /// 구역의 시작부터 (0x0040)
    Section = 0x0040,
    /// 리스트의 시작부터 (0x0050)
    List = 0x0050,
    /// 컨트롤의 시작부터 (0x0060)
    Control = 0x0060,
    /// 문서의 시작부터 (0x0070)
    Document = 0x0070,
}

/// ScanEndingPosition. `InitScan` `range` 끝 위치 (SDK: `scanEposXxx`)
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u32)]
pub enum ScanEpos {
    /// 캐럿 위치까지 (0x0000)
    Current = 0x0000,
    /// 특정 위치까지 (0x0001)
    Specified = 0x0001,
    /// 줄의 끝까지 (0x0002)
    Line = 0x0002,
    /// 문단의 끝까지 (0x0003)
    Paragraph = 0x0003,
    /// 구역의 끝까지 (0x0004)
    Section = 0x0004,
    /// 리스트의 끝까지 (0x0005)
    List = 0x0005,
    /// 컨트롤의 끝까지 (0x0006)
    Control = 0x0006,
    /// 문서의 끝까지 (0x0007)
    Document = 0x0007,
}

/// `InitScan` `range` 검색 방향 (SDK: `scanForward`, `scanBackward`)
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u32)]
pub enum ScanDirection {
    /// 정방향 (0x0000)
    Forward = 0x0000,
    /// 역방향 (0x0100)
    Backward = 0x0100,
}

/// `InitScan`의 `range` 파라미터를 타입 안전하게 구성합니다.
///
/// SDK 참고: HwpAutomation_2504.pdf § InitScan (p.22)
///
/// # 사용 예
/// ```ignore
/// // 문서 시작 → 캐럿 위치
/// ScanRange::new(ScanSpos::Document, ScanEpos::Current)
/// // 캐럿 위치 → 문서 끝 (역방향)
/// ScanRange::new(ScanSpos::Current, ScanEpos::Document).backward()
/// // 선택 블록 내
/// ScanRange::within_selection()
/// ```
#[derive(Debug, Clone, Copy)]
pub struct ScanRange(u32);

impl ScanRange {
    /// 시작 위치와 끝 위치로 정방향 검색 범위를 생성합니다.
    pub const fn new(spos: ScanSpos, epos: ScanEpos) -> Self {
        Self(spos as u32 | epos as u32)
    }

    /// 역방향 검색으로 전환합니다.
    pub const fn backward(self) -> Self {
        Self(self.0 | ScanDirection::Backward as u32)
    }

    /// 검색 범위를 선택 블록으로 제한합니다. (SDK: `scanWithinSelection` = 0x00ff)
    pub const fn within_selection() -> Self {
        Self(0x00ff)
    }

    /// 내부 `u32` 값을 반환합니다.
    pub const fn as_u32(self) -> u32 {
        self.0
    }
}

// =========================================================================
// GetText 관련 타입
// =========================================================================

/// `GetText` 반환 상태 코드
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum GetTextStatus {
    /// 텍스트 정보 없음 (SDK: 0)
    None,
    /// 리스트의 끝 (SDK: 1)
    EndOfList,
    /// 일반 텍스트 (SDK: 2)
    Normal,
    /// 다음 문단 (SDK: 3)
    NextParagraph,
    /// 제어문자 내부 진입 (SDK: 4)
    EnterControl,
    /// 제어문자 빠져나옴 (SDK: 5)
    ExitControl,
    /// 초기화 안 됨 — `InitScan` 미호출 또는 실패 (SDK: 101)
    NotInitialized,
    /// 텍스트 변환 실패 (SDK: 102)
    ConversionFailed,
    /// 알 수 없는 코드
    Unknown(i32),
}

impl GetTextStatus {
    fn from_i32(code: i32) -> Self {
        match code {
            0 => Self::None,
            1 => Self::EndOfList,
            2 => Self::Normal,
            3 => Self::NextParagraph,
            4 => Self::EnterControl,
            5 => Self::ExitControl,
            101 => Self::NotInitialized,
            102 => Self::ConversionFailed,
            n => Self::Unknown(n),
        }
    }
}

/// `KeyIndicator` 메서드가 반환하는 현재 커서 및 문서 상태 정보
#[derive(Debug, Clone, Default)]
pub struct KeyIndicatorInfo {
    /// 전체 구역 수
    pub section_count: i32,
    /// 현재 구역 번호 (1-based)
    pub section: i32,
    /// 현재 페이지 번호 (쪽 번호, 1-based)
    pub page: i32,
    /// 현재 단(Column) 번호 (1-based)
    pub column: i32,
    /// 현재 줄 번호 (문서 전체 기준 아님, 1-based)
    pub line: i32,
    /// 현재 칸 번호 (1-based)
    pub pos: i32,
    /// 수정 모드 여부 (true: 수정, false: 삽입)
    pub overwrite: bool,
    /// 현재 위치한 컨트롤의 이름 (없으면 빈 문자열)
    pub control_name: String,
}

// =========================================================================
// Clear 관련 타입
// =========================================================================

/// `Clear`의 `option` 파라미터 — 편집 중인 문서 내용의 처리 방법
///
/// SDK: `HwpSaveOption` (short type)
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(i32)]
pub enum SaveOption {
    /// 문서 내용이 변경되었을 때 사용자에게 저장할지 묻는 대화상자를 띄운다 (SDK: `hwpAskSave` = 0)
    AskSave = 0,
    /// 문서 내용을 버린다 (SDK: `hwpDiscard` = 1)
    Discard = 1,
    /// 문서가 변경된 경우 저장한다 (SDK: `hwpSaveIfDirty` = 2)
    SaveIfDirty = 2,
    /// 무조건 저장한다 (SDK: `hwpSave` = 3)
    Save = 3,
}

// =========================================================================
// SetMessageBoxMode 관련 상수
// =========================================================================

/// `SetMessageBoxMode`/`GetMessageBoxMode`의 `mode` 값 (SDK: `MB_*`)
///
/// 메시지 상자 종류마다 4비트 자리가 하나씩 있고, 그 자리에 자동으로 누를 버튼을 지정한다.
/// 여러 종류를 `|`로 조합할 수 있다. 예: `msgbox::OK_IDOK | msgbox::YESNO_IDYES`
///
/// SDK 참고: HwpAutomation_2504.pdf § GetMessageBoxMode
pub mod msgbox {
    /// 메시지 상자 종류 전체 마스크 (SDK: `MB_MASK`)
    pub const MASK: i32 = 0x00FF_FFFF;

    /// 확인(MB_OK) — 확인 (SDK: `MB_OK_IDOK`)
    pub const OK_IDOK: i32 = 0x0000_0001;
    /// 확인(MB_OK) 자리 마스크 (SDK: `MB_OK_MASK`)
    pub const OK_MASK: i32 = 0x0000_000F;

    /// 확인/취소(MB_OKCANCEL) — 확인 (SDK: `MB_OKCANCEL_IDOK`)
    pub const OKCANCEL_IDOK: i32 = 0x0000_0010;
    /// 확인/취소(MB_OKCANCEL) — 취소 (SDK: `MB_OKCANCEL_IDCANCEL`)
    pub const OKCANCEL_IDCANCEL: i32 = 0x0000_0020;
    /// 확인/취소(MB_OKCANCEL) 자리 마스크 (SDK: `MB_OKCANCEL_MASK`)
    pub const OKCANCEL_MASK: i32 = 0x0000_00F0;

    /// 종료/재시도/무시(MB_ABORTRETRYIGNORE) — 종료 (SDK: `MB_ABORTRETRYIGNORE_IDABORT`)
    pub const ABORTRETRYIGNORE_IDABORT: i32 = 0x0000_0100;
    /// 종료/재시도/무시(MB_ABORTRETRYIGNORE) — 재시도 (SDK: `MB_ABORTRETRYIGNORE_IDRETRY`)
    pub const ABORTRETRYIGNORE_IDRETRY: i32 = 0x0000_0200;
    /// 종료/재시도/무시(MB_ABORTRETRYIGNORE) — 무시 (SDK: `MB_ABORTRETRYIGNORE_IDIGNORE`)
    pub const ABORTRETRYIGNORE_IDIGNORE: i32 = 0x0000_0400;
    /// 종료/재시도/무시(MB_ABORTRETRYIGNORE) 자리 마스크 (SDK: `MB_ABORTRETRYIGNORE_MASK`)
    pub const ABORTRETRYIGNORE_MASK: i32 = 0x0000_0F00;

    /// 예/아니오/취소(MB_YESNOCANCEL) — 예 (SDK: `MB_YESNOCANCEL_IDYES`)
    pub const YESNOCANCEL_IDYES: i32 = 0x0000_1000;
    /// 예/아니오/취소(MB_YESNOCANCEL) — 아니오 (SDK: `MB_YESNOCANCEL_IDNO`)
    pub const YESNOCANCEL_IDNO: i32 = 0x0000_2000;
    /// 예/아니오/취소(MB_YESNOCANCEL) — 취소 (SDK: `MB_YESNOCANCEL_IDCANCEL`)
    pub const YESNOCANCEL_IDCANCEL: i32 = 0x0000_4000;
    /// 예/아니오/취소(MB_YESNOCANCEL) 자리 마스크 (SDK: `MB_YESNOCANCEL_MASK`)
    pub const YESNOCANCEL_MASK: i32 = 0x0000_F000;

    /// 예/아니오(MB_YESNO) — 예 (SDK: `MB_YESNO_IDYES`)
    pub const YESNO_IDYES: i32 = 0x0001_0000;
    /// 예/아니오(MB_YESNO) — 아니오 (SDK: `MB_YESNO_IDNO`)
    pub const YESNO_IDNO: i32 = 0x0002_0000;
    /// 예/아니오(MB_YESNO) 자리 마스크 (SDK: `MB_YESNO_MASK`)
    pub const YESNO_MASK: i32 = 0x000F_0000;

    /// 재시도/취소(MB_RETRYCANCEL) — 재시도 (SDK: `MB_RETRYCANCEL_IDRETRY`)
    pub const RETRYCANCEL_IDRETRY: i32 = 0x0010_0000;
    /// 재시도/취소(MB_RETRYCANCEL) — 취소 (SDK: `MB_RETRYCANCEL_IDCANCEL`)
    pub const RETRYCANCEL_IDCANCEL: i32 = 0x0020_0000;
    /// 재시도/취소(MB_RETRYCANCEL) 자리 마스크 (SDK: `MB_RETRYCANCEL_MASK`)
    pub const RETRYCANCEL_MASK: i32 = 0x00F0_0000;
}

// =========================================================================
// IHwpObject 메서드 구현
// =========================================================================

impl HwpObject {
    // ── 액션 실행기·파라미터 셋 ──

    /// `HAction` 프로퍼티 — 액션 실행기
    pub fn h_action(&self) -> crate::error::Result<crate::method::h_action::HAction> {
        self.get("HAction")
    }

    /// `HParameterSet` 프로퍼티 — 파라미터 셋 컨테이너
    pub fn h_parameter_set(&self) -> crate::error::Result<crate::pset::HParameterSet> {
        self.get("HParameterSet")
    }

    // ── 기본 액션 ──

    /// `Run` — 액션을 즉시 실행합니다.
    ///
    /// 파라미터 없이 실행 가능한 단순 액션에 사용합니다.
    /// 파라미터가 필요한 경우 `h_action().get_default / execute`를 사용합니다.
    ///
    /// SDK: `void Run(BSTR action)`
    pub fn run(&self, action: &str) -> crate::error::Result<()> {
        self.call_with("Run", vec![action.into_variant()?])
    }

    // ── 프로퍼티 ──

    /// `XHwpDocuments` 프로퍼티 — HWP 문서(도큐먼트) 컬렉션
    ///
    /// SDK: `IHwpObject.XHwpDocuments(Property)` — 도큐먼트를 관리하는 XHwpDocuments Object를 얻는다.
    pub fn documents(&self) -> crate::error::Result<crate::method::xhwp::XHwpDocuments> {
        self.get("XHwpDocuments")
    }

    /// `XHwpWindows` 프로퍼티 — HWP 창 컬렉션
    pub fn windows(&self) -> crate::error::Result<crate::method::xhwp::XHwpWindows> {
        self.get("XHwpWindows")
    }

    /// `Path` 프로퍼티 — 현재 문서의 파일 경로 (저장되지 않은 문서는 빈 문자열).
    ///
    /// SDK: `IHwpObject.Path(Property)`
    pub fn path(&self) -> crate::error::Result<String> {
        self.get("Path")
    }

    // ── 메서드 ──

    /// `GetTextFile` — 문서를 문자열로 변환하여 반환합니다.
    ///
    /// # 형식 (`format`)
    /// `GetTextFileFormat` 참조
    ///
    /// # 옵션
    /// - `save_block = true` — 선택 블록만 저장 (객체 선택 시 동작 안 함)
    pub fn get_text_file(
        &self,
        format: GetTextFileFormat,
        save_block: bool,
    ) -> crate::error::Result<String> {
        let option = if save_block { "saveblock" } else { "" };
        self.call_with(
            "GetTextFile",
            vec![format.as_str().into_variant()?, option.into_variant()?],
        )
    }

    /// `SetTextFile` — 문자열로 변환된 문서를 현재 문서에 설정합니다.
    ///
    /// SDK: `long SetTextFile(VARIANT data, BSTR format, BSTR option)`
    ///
    /// # 형식 (`format`)
    /// `GetTextFileFormat` 참조. `HWP` 형식의 `data`는 BASE64로 인코딩되어 있어야
    /// 합니다.
    ///
    /// # 옵션
    /// - `insert_file = true` — 현재 캐럿 이후에 문서를 삽입합니다.
    pub fn set_text_file(
        &self,
        data: &str,
        format: GetTextFileFormat,
        insert_file: bool,
    ) -> crate::error::Result<i32> {
        let option = if insert_file { "insertfile" } else { "" };
        self.call_with(
            "SetTextFile",
            vec![
                data.into_variant()?,
                format.as_str().into_variant()?,
                option.into_variant()?,
            ],
        )
    }

    /// `SaveAs` — 현재 문서를 지정한 경로/포맷으로 저장합니다.
    ///
    /// SDK: `boolean SaveAs(BSTR Path, BSTR Format, BSTR arg)`
    /// - `path` — 저장할 전체 파일 경로
    /// - `format` — 파일 포맷 문자열. 예: `"HWP"`, `"HWPX"`, `"HWPML2X"`, `"PDF"`,
    ///   `"HTML"`, `"UNICODE"`, `"TEXT"`
    /// - `arg` — 추가 옵션 문자열(없으면 빈 문자열)
    ///
    /// SDK가 `false`를 돌려주면(저장 실패) `ExecutionFailed` 에러를 반환합니다.
    pub fn save_as(&self, path: &str, format: &str, arg: &str) -> crate::error::Result<()> {
        let ok: bool = self.call_with(
            "SaveAs",
            vec![
                path.into_variant()?,
                format.into_variant()?,
                arg.into_variant()?,
            ],
        )?;
        ensure(ok, || format!("SaveAs({path:?}, {format:?})"))
    }

    /// `Open` — 문서 파일을 엽니다.
    ///
    /// SDK: `VARIANT_BOOL Open(BSTR filename, VARIANT Format, VARIANT arg)`
    /// - `path` — 문서 파일의 경로(URL 사용 가능)
    /// - `format` — 문서 형식. 예: `"HWP"`, `"HWPML2X"`, `"HTML"`, `"TEXT"`, `"UNICODE"`.
    ///   빈 문자열이면 한글이 형식을 자동으로 인식합니다. 쓸 수 있는 형식은 설치된
    ///   문서 필터(`*.dft`)에 따라 다릅니다.
    /// - `arg` — 세부 옵션. `key:value;key:value;...` 형식이고 의미는 `format`에 따라
    ///   다릅니다. 예: `"forceopen:true;suspendpassword:true"`. 없으면 빈 문자열.
    ///
    /// SDK가 `false`를 돌려주면(열기 실패) `ExecutionFailed` 에러를 반환합니다.
    pub fn open(&self, path: &str, format: &str, arg: &str) -> crate::error::Result<()> {
        let ok: bool = self.call_with(
            "Open",
            vec![
                path.into_variant()?,
                format.into_variant()?,
                arg.into_variant()?,
            ],
        )?;
        ensure(ok, || format!("Open({path:?}, {format:?}, {arg:?})"))
    }

    /// `Clear` — 현재 편집 중인 문서의 내용을 닫고 빈 문서 편집 상태로 돌아갑니다.
    ///
    /// SDK: `void Clear([HwpSaveOption option])`
    /// - `option` — 편집 중인 문서 내용의 처리 방법. [`SaveOption`] 참조.
    ///
    /// [`SaveOption::SaveIfDirty`]나 [`SaveOption::Save`]를 주었는데 현재 문서 경로가 없으면
    /// 한글이 "새 이름으로 저장" 대화상자를 띄웁니다.
    pub fn clear(&self, option: SaveOption) -> crate::error::Result<()> {
        self.call_with("Clear", vec![(option as i32).into_variant()?])
    }

    /// `Quit` — 한글을 종료합니다.
    ///
    /// 호출한 뒤에는 이 `HwpObject`와 그 복제본을 더 쓸 수 없습니다.
    ///
    /// SDK: `void Quit()` (IDL에만 있고 `HwpAutomation_2504.pdf`에는 설명이 없음)
    pub fn quit(&self) -> crate::error::Result<()> {
        self.call("Quit")
    }

    /// `GetMessageBoxMode` — 현재 메시지 상자 mode를 얻습니다.
    ///
    /// SDK: `long GetMessageBoxMode()`
    /// - 반환값 — [`msgbox`]의 상수를 조합한 값
    pub fn get_message_box_mode(&self) -> crate::error::Result<i32> {
        self.call("GetMessageBoxMode")
    }

    /// `SetMessageBoxMode` — 한글의 메시지 상자를 띄우지 않고, 지정한 버튼을 자동으로
    /// 누른 효과를 줍니다.
    ///
    /// SDK: `long SetMessageBoxMode(long mode)`
    /// - `mode` — [`msgbox`]의 상수를 `|`로 조합한 값
    ///
    /// 반환값은 이 호출 전에 적용되어 있던 mode입니다. 되돌릴 때 다시 넘기면 됩니다.
    pub fn set_message_box_mode(&self, mode: i32) -> crate::error::Result<i32> {
        self.call_with("SetMessageBoxMode", vec![mode.into_variant()?])
    }

    /// `RegisterModule` — 보안 모듈을 등록합니다.
    ///
    /// SDK: `VARIANT_BOOL RegisterModule(BSTR ModuleType, VARIANT ModuleData)`
    ///
    /// 파일 경로 검사 모듈을 등록하면, 한글이 파일에 접근할 때 띄우는 접근 허용 대화상자 대신
    /// 그 모듈 DLL에 허용 여부를 묻습니다: `register_module("FilePathCheckDLL", <모듈 이름>)`
    ///
    /// - 사전 조건: 레지스트리 `HKEY_CURRENT_USER\Software\HNC\HwpAutomation\Modules`에
    ///   `<모듈 이름>`을 이름으로 하는 문자열(REG_SZ) 값으로 DLL의 전체 경로를 써 둡니다.
    /// - DLL은 `BOOL __stdcall IsAccessiblePath(HWND hWnd, LONG nID, LPCTSTR szFilePath,
    ///   LPCTSTR szSiteInfo)`를 export하고, 접근을 허용하면 `TRUE`를 돌려줍니다.
    ///   문자열이 ANSI인지 UTF-16인지는 확인되지 않았습니다.
    ///
    /// 근거: [devcenter-archive](https://github.com/hancom-io/devcenter-archive)의
    /// `hwp-automation/보안모듈(Automation).zip` (예제 DLL 소스와 레지스트리 화면).
    /// `HwpAutomation_2504.pdf`의 설명은 "보안 모듈 관련 문서 참조"뿐입니다.
    ///
    /// SDK가 `false`를 돌려주면(등록 실패) `ExecutionFailed` 에러를 반환합니다.
    pub fn register_module(
        &self,
        module_type: &str,
        module_data: &str,
    ) -> crate::error::Result<()> {
        let ok: bool = self.call_with(
            "RegisterModule",
            vec![module_type.into_variant()?, module_data.into_variant()?],
        )?;
        ensure(ok, || format!("RegisterModule({module_type:?}, {module_data:?})"))
    }

    /// `InitScan` — 문서 텍스트 검색을 초기화합니다.
    ///
    /// `GetText`를 호출하기 전에 반드시 호출해야 합니다.
    /// 사용이 끝나면 `ReleaseScan`으로 정리해야 합니다.
    ///
    /// # 매개변수
    /// - `option` — 검색 대상 (`mask::*` 조합)
    /// - `range` — 검색 범위 (`ScanRange`)
    /// - `spara`, `spos` — 시작 문단/위치 (`ScanSpos::Specified` 사용 시)
    /// - `epara`, `epos` — 끝 문단/위치 (`ScanEpos::Specified` 사용 시)
    ///
    /// SDK가 `false`를 돌려주면 `ExecutionFailed` 에러를 반환합니다.
    pub fn init_scan(
        &self,
        option: u32,
        range: ScanRange,
        spara: u32,
        spos: u32,
        epara: u32,
        epos: u32,
    ) -> crate::error::Result<()> {
        let range_val = range.as_u32();
        tracing::trace!(
            option = format_args!("0x{option:02x}"),
            range = format_args!("0x{range_val:04x}"),
            spara,
            spos,
            epara,
            epos,
            "InitScan"
        );
        let ok: bool = self.call_with(
            "InitScan",
            vec![
                option.into_variant()?,
                range_val.into_variant()?,
                spara.into_variant()?,
                spos.into_variant()?,
                epara.into_variant()?,
                epos.into_variant()?,
            ],
        )?;
        ensure(ok, || format!("InitScan(option=0x{option:02x}, range=0x{range_val:04x})"))
    }

    /// `GetText` — 커서가 가리키는 단락의 텍스트를 읽고, 커서를 다음 단락으로 이동합니다.
    ///
    /// `InitScan`으로 스캔을 초기화한 뒤 반복 호출하면 문서를 단락 단위로 순서대로 읽을 수 있습니다.
    /// `GetTextStatus`가 `EndOfDocument`이면 더 이상 읽을 내용이 없습니다.
    ///
    /// 반환: `(GetTextStatus, 텍스트)`
    pub fn get_text(&self) -> crate::error::Result<(GetTextStatus, String)> {
        let (result_var, text) =
            com_util::call_method_with_bstr_out(&self.dispatch, "GetText", vec![])?;
        let code = i32::from_variant(&result_var)?;
        let status = GetTextStatus::from_i32(code);
        tracing::trace!(code, ?status, len = text.len(), "GetText");
        Ok((status, text))
    }

    /// `ReleaseScan` — `InitScan`으로 설정된 검색 상태를 해제합니다.
    pub fn release_scan(&self) -> crate::error::Result<()> {
        self.call("ReleaseScan")
    }

    /// `GetPos` — 현재 캐럿 위치를 `(list, para, pos)`로 반환합니다.
    ///
    /// - `list` : 캐럿이 위치한 리스트 아이디
    /// - `para` : 문단 번호
    /// - `pos`  : 문단 내 문자 단위 위치
    ///
    /// 반환된 triple은 [`select_text`](Self::select_text)나 SDK의 `SetPos`에
    /// 그대로 넘길 수 있습니다.
    ///
    /// SDK: `void GetPos(long FAR* list, long FAR* para, long FAR* pos)`
    pub fn get_pos(&self) -> crate::error::Result<(i32, i32, i32)> {
        use windows::Win32::System::Variant::{VARIANT, VT_BYREF, VT_I4};

        let mut list = 0i32;
        let mut para = 0i32;
        let mut pos = 0i32;

        let args = unsafe {
            let mut v_list = VARIANT::default();
            (*v_list.Anonymous.Anonymous).vt = VT_I4 | VT_BYREF;
            (*v_list.Anonymous.Anonymous).Anonymous.plVal = &mut list;

            let mut v_para = VARIANT::default();
            (*v_para.Anonymous.Anonymous).vt = VT_I4 | VT_BYREF;
            (*v_para.Anonymous.Anonymous).Anonymous.plVal = &mut para;

            let mut v_pos = VARIANT::default();
            (*v_pos.Anonymous.Anonymous).vt = VT_I4 | VT_BYREF;
            (*v_pos.Anonymous.Anonymous).Anonymous.plVal = &mut pos;

            vec![v_list, v_para, v_pos]
        };

        let _ = com_util::call_method_with(&self.dispatch, "GetPos", args)?;
        Ok((list, para, pos))
    }

    /// `SelectText` — 문단/위치 쌍으로 표현된 범위를 직접 선택합니다.
    ///
    /// HAction의 `Select`/`MoveSel*` 경로가 IME 조합 상태에서 동작하지 않는
    /// 경우의 우회로로 유용합니다. Automation 경로라 키/포커스 상태에 덜
    /// 민감합니다.
    ///
    /// SDK: `BOOL SelectText(long spara, long spos, long epara, long epos)`
    ///
    /// SDK가 `false`를 돌려주면 `ExecutionFailed` 에러를 반환합니다.
    pub fn select_text(
        &self,
        spara: i32,
        spos: i32,
        epara: i32,
        epos: i32,
    ) -> crate::error::Result<()> {
        let ok: bool = self.call_with(
            "SelectText",
            vec![
                spara.into_variant()?,
                spos.into_variant()?,
                epara.into_variant()?,
                epos.into_variant()?,
            ],
        )?;
        ensure(ok, || format!("SelectText({spara}, {spos}, {epara}, {epos})"))
    }

    /// `KeyIndicator` — 현재 커서 위치의 정보를 가져오고 상태 표시줄을 강제로 갱신합니다.
    ///
    /// 액션 실행 후 커서 위치가 화면과 일치하지 않을 때 호출하면 내부 상태를 동기화할 수 있습니다.
    ///
    /// SDK: `BOOL KeyIndicator(long* seccnt, long* secno, long* prnpageno, long* colno, long* line, long* pos, short* over, BSTR* ctrlname)`
    pub fn key_indicator(&self) -> crate::error::Result<KeyIndicatorInfo> {
        use windows::Win32::System::Variant::{VARIANT, VT_BSTR, VT_BYREF, VT_I2, VT_I4};

        // 출력 인자를 위한 지역 변수
        let mut seccnt = 0i32;
        let mut secno = 0i32;
        let mut prnpageno = 0i32;
        let mut colno = 0i32;
        let mut line = 0i32;
        let mut pos = 0i32;
        let mut over = 0i16;
        let mut ctrlname = windows::core::BSTR::new();

        // VT_BYREF VARIANT 직접 구성
        let args = unsafe {
            let mut v_seccnt = VARIANT::default();
            (*v_seccnt.Anonymous.Anonymous).vt = VT_I4 | VT_BYREF;
            (*v_seccnt.Anonymous.Anonymous).Anonymous.plVal = &mut seccnt;

            let mut v_secno = VARIANT::default();
            (*v_secno.Anonymous.Anonymous).vt = VT_I4 | VT_BYREF;
            (*v_secno.Anonymous.Anonymous).Anonymous.plVal = &mut secno;

            let mut v_prnpageno = VARIANT::default();
            (*v_prnpageno.Anonymous.Anonymous).vt = VT_I4 | VT_BYREF;
            (*v_prnpageno.Anonymous.Anonymous).Anonymous.plVal = &mut prnpageno;

            let mut v_colno = VARIANT::default();
            (*v_colno.Anonymous.Anonymous).vt = VT_I4 | VT_BYREF;
            (*v_colno.Anonymous.Anonymous).Anonymous.plVal = &mut colno;

            let mut v_line = VARIANT::default();
            (*v_line.Anonymous.Anonymous).vt = VT_I4 | VT_BYREF;
            (*v_line.Anonymous.Anonymous).Anonymous.plVal = &mut line;

            let mut v_pos = VARIANT::default();
            (*v_pos.Anonymous.Anonymous).vt = VT_I4 | VT_BYREF;
            (*v_pos.Anonymous.Anonymous).Anonymous.plVal = &mut pos;

            let mut v_over = VARIANT::default();
            (*v_over.Anonymous.Anonymous).vt = VT_I2 | VT_BYREF;
            (*v_over.Anonymous.Anonymous).Anonymous.piVal = &mut over;

            let mut v_ctrlname = VARIANT::default();
            (*v_ctrlname.Anonymous.Anonymous).vt = VT_BSTR | VT_BYREF;
            (*v_ctrlname.Anonymous.Anonymous).Anonymous.pbstrVal =
                &mut ctrlname as *mut windows::core::BSTR;

            vec![
                v_seccnt,
                v_secno,
                v_prnpageno,
                v_colno,
                v_line,
                v_pos,
                v_over,
                v_ctrlname,
            ]
        };

        // com_util::call_method_with를 통해 호출
        let res_var = com_util::call_method_with(&self.dispatch, "KeyIndicator", args)?;
        let success = bool::from_variant(&res_var)?;

        ensure(success, || "KeyIndicator()".to_string())?;

        Ok(KeyIndicatorInfo {
            section_count: seccnt,
            section: secno,
            page: prnpageno,
            column: colno,
            line,
            pos,
            overwrite: over != 0,
            control_name: ctrlname.to_string(),
        })
    }
}
