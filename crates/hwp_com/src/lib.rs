/// 한글(HWP) OLE 클라이언트 라이브러리
///
/// 별도 프로세스에서 `CoCreateInstance` 또는 Running Object Table을 통해
/// 한글을 제어합니다. ProgID: `"HwpFrame.HwpObject.2"`
///
/// # 사용 예
/// ```ignore
/// use hwp_com::HwpClient;
///
/// let hwp = HwpClient::new()?;       // 새 한글 인스턴스 시작
/// hwp.run("FileNew")?;               // 새 문서
/// hwp.windows()?.active()?.set_visible(true)?; // 창 표시
/// ```
pub use hwp_core::error::{HwpError, Result};
use hwp_core::hwp_obj::HwpObject;
use windows::Win32::System::Com::{
    CLSCTX_LOCAL_SERVER, CLSIDFromProgID, CoCreateInstance, CreateBindCtx, GetRunningObjectTable,
    IDispatch,
};
use windows::Win32::System::Ole::OleInitialize;
use windows::core::{Interface, w};

/// OLE 초기화 후 `HwpObject`를 생성/연결하는 클라이언트
pub struct HwpClient;

impl HwpClient {
    /// 새 한글 인스턴스를 시작하고 `HwpObject`를 반환합니다.
    ///
    /// 내부적으로 `OleInitialize` → `CLSIDFromProgID` → `CoCreateInstance`를 호출합니다.
    pub fn new() -> Result<HwpObject> {
        Self::ole_init()?;
        let dispatch = Self::create_instance()?;
        HwpObject::new(dispatch)
    }

    /// 이미 실행 중인 한글 인스턴스에 연결합니다.
    ///
    /// Running Object Table에서 `"!IHwpObject.130"` 등의 모니커를 검색합니다.
    /// 여러 인스턴스가 실행 중이면 첫 번째로 발견된 인스턴스에 연결합니다.
    pub fn attach() -> Result<HwpObject> {
        Self::ole_init()?;
        let dispatch = Self::attach_running()?;
        HwpObject::new(dispatch)
    }

    /// 실행 중인 모든 한글 인스턴스를 열거하여 `(ROT 모니커 이름, HwpObject)` 쌍으로 반환합니다.
    ///
    /// Running Object Table 전체를 훑어 `"!IHwpObject"`를 포함하는 모든 모니커를 수집합니다.
    /// 열린 한글이 없으면 빈 벡터를 반환합니다.
    pub fn list_running() -> Result<Vec<(String, HwpObject)>> {
        Self::ole_init()?;
        let entries = Self::collect_running()?;
        let mut out = Vec::with_capacity(entries.len());
        for (name, dispatch) in entries {
            out.push((name, HwpObject::new(dispatch)?));
        }
        Ok(out)
    }

    /// OLE 런타임을 초기화합니다.
    fn ole_init() -> Result<()> {
        unsafe {
            OleInitialize(None).map_err(|_| HwpError::OleInitFailed)?;
        }
        Ok(())
    }

    /// `CoCreateInstance`로 새 HWP 프로세스를 생성합니다.
    fn create_instance() -> Result<IDispatch> {
        unsafe {
            let clsid = CLSIDFromProgID(w!("HwpFrame.HwpObject.2"))
                .map_err(|_| HwpError::ConnectionFailed)?;
            let dispatch: IDispatch = CoCreateInstance(&clsid, None, CLSCTX_LOCAL_SERVER)
                .map_err(|_| HwpError::ConnectionFailed)?;
            Ok(dispatch)
        }
    }

    /// Running Object Table에서 실행 중인 한글 인스턴스를 검색합니다.
    fn attach_running() -> Result<IDispatch> {
        Self::collect_running()?
            .into_iter()
            .next()
            .map(|(_, d)| d)
            .ok_or(HwpError::ConnectionFailed)
    }

    /// ROT를 순회하며 `"!IHwpObject"`가 포함된 모든 모니커를 `(이름, IDispatch)`로 수집합니다.
    fn collect_running() -> Result<Vec<(String, IDispatch)>> {
        unsafe {
            let rot = GetRunningObjectTable(0).map_err(|_| HwpError::ConnectionFailed)?;
            let enum_moniker = rot.EnumRunning().map_err(|_| HwpError::ConnectionFailed)?;
            let bind_ctx = CreateBindCtx(0).map_err(|_| HwpError::ConnectionFailed)?;

            let target = "!HwpObject";
            let mut results = Vec::new();

            loop {
                let mut moniker = [None];
                let mut fetched = 0u32;
                let hr = enum_moniker.Next(&mut moniker, Some(&mut fetched));
                if hr.is_err() || fetched == 0 {
                    break;
                }

                let moniker = match &moniker[0] {
                    Some(m) => m,
                    None => continue,
                };

                let display_name = moniker
                    .GetDisplayName(&bind_ctx, None)
                    .map_err(|_| HwpError::ConnectionFailed)?;
                let name = display_name.to_string().unwrap_or_default();

                if name.contains(target) {
                    let unknown = rot
                        .GetObject(moniker)
                        .map_err(|_| HwpError::ConnectionFailed)?;
                    let dispatch: IDispatch =
                        unknown.cast().map_err(|_| HwpError::ConnectionFailed)?;
                    results.push((name, dispatch));
                }
            }

            Ok(results)
        }
    }
}
