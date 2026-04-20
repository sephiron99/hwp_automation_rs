/// SDK `IHwpObject` — 한글 OLE Automation 최상위 객체의 메서드 구현
///
/// SDK 참고: HwpAutomation_2504.pdf § IHwpObject
use crate::com_util;
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
// IHwpObject 메서드 구현
// =========================================================================

impl HwpObject {
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
    pub fn documents(&self) -> crate::error::Result<crate::hwp_types::XHwpDocuments> {
        self.get("XHwpDocuments")
    }

    /// `XHwpWindows` 프로퍼티 — HWP 창 컬렉션
    pub fn windows(&self) -> crate::error::Result<crate::hwp_types::XHwpWindows> {
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
    pub fn init_scan(
        &self,
        option: u32,
        range: ScanRange,
        spara: u32,
        spos: u32,
        epara: u32,
        epos: u32,
    ) -> crate::error::Result<bool> {
        let range_val = range.as_u32();
        crate::debug::log(
            "InitScan",
            &format!("option=0x{option:02x}, range=0x{range_val:04x}, spara={spara}, spos={spos}, epara={epara}, epos={epos}"),
        );
        self.call_with(
            "InitScan",
            vec![
                option.into_variant()?,
                range_val.into_variant()?,
                spara.into_variant()?,
                spos.into_variant()?,
                epara.into_variant()?,
                epos.into_variant()?,
            ],
        )
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
        // crate::debug::log(
        //     "GetText",
        //     &format!("code={code}, status={status:?}, len={}, text={text:?}", text.len()),
        // );
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
    pub fn select_text(
        &self,
        spara: i32,
        spos: i32,
        epara: i32,
        epos: i32,
    ) -> crate::error::Result<bool> {
        self.call_with(
            "SelectText",
            vec![
                spara.into_variant()?,
                spos.into_variant()?,
                epara.into_variant()?,
                epos.into_variant()?,
            ],
        )
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

        if !success {
            return Err(crate::error::HwpError::ComError(windows::core::Error::new(
                windows::core::HRESULT(-1),
                "KeyIndicator failed",
            )));
        }

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
