use crate::com_util;
use crate::hwp_ver::HwpVer;
use crate::variant::FromVariant;
use windows::core::Interface;
use windows::Win32::System::Com::IDispatch;
use windows::Win32::System::Variant::VARIANT;

/// 한글 OLE Automation 최상위 객체 래퍼 (`IHwpObject` 대응)
pub struct HwpObject {
    pub(crate) dispatch: IDispatch,
    ver: HwpVer,
}

impl HwpObject {
    fn detect_version(dispatch: &IDispatch) -> HwpVer {
        let var = match com_util::get_property(dispatch, "Version") {
            Ok(v) => v,
            Err(_) => return HwpVer::Other(0, 0),
        };
        // OLE 클라이언트: "10, 0, 0, 14727" 같은 문자열
        if let Ok(s) = String::from_variant(&var) {
            return HwpVer::from_version_string(&s);
        }
        // Add-in DLL: i32 코드
        if let Ok(code) = i32::from_variant(&var) {
            return HwpVer::from_u32(code as u32);
        }
        HwpVer::Other(0, 0)
    }

    /// 소유권 있는 `IDispatch`로 `HwpObject`를 생성합니다.
    pub fn new(dispatch: IDispatch) -> crate::error::Result<Self> {
        let ver = Self::detect_version(&dispatch);
        Ok(Self { dispatch, ver })
    }

    /// Addon FFI용: raw `IDispatch` 포인터에서 `HwpObject`를 생성합니다.
    ///
    /// # Safety
    /// HWP가 소유한 포인터이므로 `AddRef`/`Release`를 호출하면 안 됩니다.
    /// 호출자는 반드시 `std::mem::forget`으로 Drop을 방지해야 합니다.
    pub unsafe fn from_raw_dispatch(raw: *mut std::ffi::c_void) -> crate::error::Result<Self> {
        if raw.is_null() {
            return Err(crate::error::HwpError::ExecutionFailed(
                "null IDispatch pointer".to_string(),
            ));
        }
        let dispatch: IDispatch = unsafe { IDispatch::from_raw(raw) };
        let ver = Self::detect_version(&dispatch);
        Ok(Self { dispatch, ver })
    }

    /// 현재 한글 버전을 반환합니다.
    pub fn version(&self) -> &HwpVer {
        &self.ver
    }

    /// COM 마샬링을 위해 내부 `IDispatch`의 참조를 반환합니다.
    pub fn as_dispatch(&self) -> &IDispatch {
        &self.dispatch
    }

    /// `IDispatch`의 raw 포인터를 반환합니다.
    ///
    /// `from_raw_dispatch`의 역변환입니다.
    /// 참조 카운트에 영향을 주지 않으므로, 반환된 포인터로 `HwpObject`를 생성한 경우
    /// 반드시 `std::mem::forget`으로 Drop을 방지해야 합니다.
    pub fn as_raw_dispatch(&self) -> *mut std::ffi::c_void {
        self.dispatch.as_raw()
    }

    // ── 범용 COM 접근 헬퍼 ──

    /// 프로퍼티 읽기
    pub fn get<T: FromVariant>(&self, name: &str) -> crate::error::Result<T> {
        let var = com_util::get_property(&self.dispatch, name)?;
        T::from_variant(&var)
    }

    /// 인자 없는 메서드 호출
    pub fn call<T: FromVariant>(&self, name: &str) -> crate::error::Result<T> {
        let var = com_util::call_method(&self.dispatch, name)?;
        T::from_variant(&var)
    }

    /// 인자 있는 메서드 호출
    pub fn call_with<T: FromVariant>(
        &self,
        name: &str,
        args: Vec<VARIANT>,
    ) -> crate::error::Result<T> {
        let var = com_util::call_method_with(&self.dispatch, name, args)?;
        T::from_variant(&var)
    }
}
