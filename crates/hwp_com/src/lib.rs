//! 한글(HWP) OLE 클라이언트 라이브러리
//!
//! 별도 프로세스에서 `CoCreateInstance` 또는 Running Object Table을 통해
//! 한글을 제어합니다. ProgID: `"HwpFrame.HwpObject.2"`
//!
//! OLE는 [`HwpClient`]가 스레드마다 처음 한 번 초기화하고, 그 스레드가 끝날 때 해제합니다.
//! 해제 시점을 직접 정하려면 [`OleGuard`]를 쓰면 됩니다.
//!
//! # 사용 예
//! ```ignore
//! use hwp_com::HwpClient;
//!
//! let hwp = HwpClient::new()?;       // 이 스레드의 OLE 초기화 + 새 한글 인스턴스 시작
//! hwp.run("FileNew")?;               // 새 문서
//! hwp.windows()?.active()?.set_visible(true)?; // 창 표시
//! ```

#![cfg_attr(docsrs, feature(doc_cfg))]

// `hwp_core`에 따로 의존하지 않아도 되도록 사용자에게 필요한 타입과 층 모듈을 re-export한다.
// 내부용 모듈(`com_util`, `disp_obj`, `hwp_obj`)은 re-export에서 제외한다.
// 같은 항목이므로 `hwp_core::X`와 섞어 써도 된다.
#[doc(no_inline)]
pub use hwp_core::HwpObject;
#[doc(no_inline)]
pub use hwp_core::error::{HwpError, Result};
#[doc(no_inline)]
pub use hwp_core::hwp_ver::HwpVer;
#[doc(no_inline)]
pub use hwp_core::{actions, error, ext, hwp_ver, method, pos, pset, variant};

use windows::Win32::Foundation::CO_E_NOTINITIALIZED;
use windows::Win32::System::Com::{
    CLSCTX_LOCAL_SERVER, CLSIDFromProgID, CoCreateInstance, CreateBindCtx, GetRunningObjectTable,
    IDispatch,
};
use windows::Win32::System::Ole::{OleInitialize, OleUninitialize};
use windows::core::{Interface, w};

/// 현재 스레드의 OLE 초기화. drop되면 `OleUninitialize`를 부릅니다.
///
/// `OleInitialize`는 스레드마다 따로 부르고, 성공한 호출마다 같은 스레드에서
/// `OleUninitialize`를 한 번씩 불러야 합니다. 이 guard는 그 짝을 맞추며, 다른 스레드로
/// 옮길 수 없습니다(`!Send`, `!Sync`).
///
/// [`HwpClient`]가 스레드마다 하나씩 자동으로 만들어 두므로 보통은 직접 쓸 필요가 없습니다.
/// OLE를 스레드 종료보다 먼저 해제하고 싶을 때 씁니다. 자동 guard와 함께 있어도 초기화 횟수가
/// 하나 늘 뿐이라 짝은 맞지만, 자동 guard가 남아 있는 동안은 OLE가 실제로 해제되지 않습니다.
///
/// guard가 drop되기 전에 그 스레드의 모든 [`HwpObject`]가 drop되어야 합니다. 같은 스코프라면
/// guard를 먼저 선언하면 됩니다. 이미 OLE를 초기화한 스레드에서 다시 만들어도 되고, 그때는
/// 초기화 횟수만 하나 늘었다가 drop될 때 줄어듭니다.
///
/// 이 스레드가 이미 멀티스레드 아파트(MTA)로 초기화돼 있으면 `RPC_E_CHANGED_MODE`로
/// 실패합니다(`HwpError::ComError`).
#[must_use = "drop되는 즉시 OLE가 해제됩니다"]
pub struct OleGuard(std::marker::PhantomData<*const ()>);

impl OleGuard {
    /// 현재 스레드에서 `OleInitialize`를 부릅니다.
    pub fn new() -> Result<Self> {
        unsafe { OleInitialize(None) }?;
        Ok(Self(std::marker::PhantomData))
    }
}

impl Drop for OleGuard {
    fn drop(&mut self) {
        unsafe { OleUninitialize() };
    }
}

thread_local! {
    /// [`HwpClient`]가 이 스레드에서 처음 불릴 때 만드는 OLE 초기화. 스레드가 끝날 때 drop된다.
    static THREAD_OLE: std::cell::OnceCell<OleGuard> = const { std::cell::OnceCell::new() };
}

/// 이 스레드에 자동 OLE guard가 없으면 만든다. 이미 있으면 아무것도 하지 않는다.
fn ensure_thread_ole() -> Result<()> {
    THREAD_OLE.with(|cell| {
        if cell.get().is_none() {
            // 실패하면 저장하지 않으므로 다음 호출에서 다시 시도한다.
            let _ = cell.set(OleGuard::new()?);
        }
        Ok(())
    })
}

/// 실행 중인 한글을 새로 띄우거나 찾아 [`HwpObject`]를 얻는 OLE 클라이언트
///
/// SDK `HwpAutomation_2504.pdf` 1.4절(외부 클라이언트와 글)과 1.5절(외부 클라이언트와 실행 중인
/// 글)의 접속 방법에 대응합니다.
///
/// # OLE 초기화
/// 각 함수는 그 스레드에서 처음 불릴 때 OLE를 초기화하고([`OleGuard`]를 thread-local에 둠),
/// 같은 스레드에서 다시 불리면 그것을 그대로 씁니다. 해제(`OleUninitialize`)는 스레드가 끝날 때
/// 일어납니다. 메인 스레드에서는 Rust가 thread-local 정리를 보장하지 않지만, 그때는 프로세스가
/// 끝나므로 운영체제가 정리합니다.
///
/// thread-local끼리의 drop 순서는 정해져 있지 않으므로, 얻은 [`HwpObject`]나 그 하위 객체를
/// `static`이나 다른 thread-local에 두지 마세요. OLE가 먼저 해제된 뒤에 COM 참조가 풀릴 수
/// 있습니다. 지역 변수나 스레드 안에서 끝나는 값이면 문제없습니다.
///
/// 이 스레드가 이미 멀티스레드 아파트(MTA)로 초기화돼 있으면 `RPC_E_CHANGED_MODE`가 담긴
/// `HwpError::ComError`로 실패합니다.
pub struct HwpClient;

impl HwpClient {
    /// 새 한글 인스턴스를 시작하고 `HwpObject`를 반환합니다.
    ///
    /// 내부적으로 (이 스레드에서 처음이면 `OleInitialize` →) `CLSIDFromProgID` → `CoCreateInstance`를 호출합니다.
    pub fn new() -> Result<HwpObject> {
        ensure_thread_ole()?;
        let dispatch = Self::create_instance()?;
        HwpObject::new(dispatch)
    }

    /// 이미 실행 중인 한글 인스턴스에 연결합니다.
    ///
    /// Running Object Table에서 `"!IHwpObject.130"` 등의 모니커를 검색합니다.
    /// 여러 인스턴스가 실행 중이면 첫 번째로 발견된 인스턴스에 연결합니다.
    pub fn attach() -> Result<HwpObject> {
        ensure_thread_ole()?;
        let dispatch = Self::attach_running()?;
        HwpObject::new(dispatch)
    }

    /// 실행 중인 모든 한글 인스턴스를 열거하여 `(ROT 모니커 이름, HwpObject)` 쌍으로 반환합니다.
    ///
    /// Running Object Table 전체를 훑어 `"!IHwpObject"`를 포함하는 모든 모니커를 수집합니다.
    /// 열린 한글이 없으면 빈 벡터를 반환합니다.
    pub fn list_running() -> Result<Vec<(String, HwpObject)>> {
        ensure_thread_ole()?;
        let entries = Self::collect_running()?;
        let mut out = Vec::with_capacity(entries.len());
        for (name, dispatch) in entries {
            out.push((name, HwpObject::new(dispatch)?));
        }
        Ok(out)
    }

    /// `CoCreateInstance`로 새 HWP 프로세스를 생성합니다.
    fn create_instance() -> Result<IDispatch> {
        unsafe {
            let clsid = CLSIDFromProgID(w!("HwpFrame.HwpObject.2")).map_err(connection_err)?;
            let dispatch: IDispatch =
                CoCreateInstance(&clsid, None, CLSCTX_LOCAL_SERVER).map_err(connection_err)?;
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
            let rot = GetRunningObjectTable(0).map_err(connection_err)?;
            let enum_moniker = rot.EnumRunning().map_err(connection_err)?;
            let bind_ctx = CreateBindCtx(0).map_err(connection_err)?;

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
                    .map_err(connection_err)?;
                let name = display_name.to_string().unwrap_or_default();

                if name.contains(target) {
                    let unknown = rot.GetObject(moniker).map_err(connection_err)?;
                    let dispatch: IDispatch = unknown.cast().map_err(connection_err)?;
                    results.push((name, dispatch));
                }
            }

            Ok(results)
        }
    }
}

/// 접속 단계의 COM 에러를 바꾼다. OLE를 초기화하지 않은 경우(`CO_E_NOTINITIALIZED`)는
/// 한글 설치 문제가 아니므로 원래 에러를 그대로 돌려준다.
fn connection_err(e: windows::core::Error) -> HwpError {
    if e.code() == CO_E_NOTINITIALIZED {
        HwpError::ComError(e)
    } else {
        HwpError::ConnectionFailed
    }
}
