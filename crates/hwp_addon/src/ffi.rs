use crate::hwp_user_action::HwpUserAction;
use hwp_core::hwp_obj::HwpObject;
use std::ffi::{CString, c_char, c_void};
use std::sync::OnceLock;

// =========================================================================
// C++ IHncUserActionModule vtable 호환 레이아웃
//
// C++ 원본 (UserActionModule.h):
//   interface IHncUserActionModule {
//       virtual LPCSTR EnumAction(int) = 0;
//       virtual BOOL GetActionImage(LPCSTR, UINT, HBITMAP*, int*) = 0;
//       virtual BOOL UpdateUI(LPCSTR, LPDISPATCH, UINT*) = 0;
//       virtual int DoAction(LPCSTR, LPDISPATCH) = 0;
//   };
//
// ── __thiscall 호출 규약 ──
// MSVC 32-bit(x86)에서 C++ 가상 함수는 __thiscall 규약을 사용한다:
//   - this 포인터 → ECX 레지스터로 전달 (스택이 아님)
//   - 나머지 인자 → 오른쪽에서 왼쪽 순으로 스택에 push
//   - 스택 정리 → callee(함수 쪽)가 담당
//
// Rust의 기본 extern "C" / "system"은 this를 포함한 모든 인자를 스택으로
// 전달하므로, 규약이 맞지 않아 인자가 깨진다.
// extern "thiscall"을 지정해야 HWP가 ECX에 넣은 this를 올바르게 받을 수 있다.
// 잘못 지정하면 extract_plugin이 쓰레기 포인터를 역참조해 크래시 발생.
//
// LPCSTR = const char* (ANSI, NOT Unicode)
// =========================================================================

#[repr(C)]
#[allow(non_snake_case)]
pub struct IHncUserActionModuleVtbl {
    pub EnumAction: unsafe extern "thiscall" fn(this: *mut c_void, n: i32) -> *const c_char,
    pub GetActionImage: unsafe extern "thiscall" fn(
        this: *mut c_void,
        sz_action: *const c_char,
        u_state: u32,
        ph_bitmap: *mut isize,
        pn_image_index: *mut i32,
    ) -> i32,
    pub UpdateUI: unsafe extern "thiscall" fn(
        this: *mut c_void,
        sz_action: *const c_char,
        p_obj: *mut c_void,
        lpu_state: *mut u32,
    ) -> i32,
    pub DoAction: unsafe extern "thiscall" fn(
        this: *mut c_void,
        sz_action: *const c_char,
        p_obj: *mut c_void,
    ) -> i32,
}

/// Rust 플러그인을 감싸는 C++ 호환 객체
///
/// 메모리 레이아웃:
/// ```text
/// [0..4]  lpVtbl → IHncUserActionModuleVtbl (C++ vptr 위치)
/// [4..]   plugin  (사용자 구현 HwpUserAction)
/// [..]    action_names (EnumAction 반환용 C 문자열 캐시)
/// ```
#[repr(C)]
#[allow(non_snake_case)]
pub struct RustActionModule<T: HwpUserAction> {
    pub lpVtbl: *const IHncUserActionModuleVtbl,
    pub plugin: T,
    /// `EnumAction`이 돌려줄 NUL 종료 액션 이름. 첫 `EnumAction` 호출 때 한 번
    /// 채우고 이후 바꾸지 않는다. 모듈이 `static`이므로 각 `CString`의 주소는
    /// DLL이 로드된 동안 유효하고 내용도 그대로다.
    action_names: OnceLock<Box<[CString]>>,
}

impl<T: HwpUserAction> RustActionModule<T> {
    /// `export_hwp_addon!` 매크로가 `static MODULE`을 만들 때 쓰는 생성자.
    pub const fn new(lp_vtbl: *const IHncUserActionModuleVtbl, plugin: T) -> Self {
        Self {
            lpVtbl: lp_vtbl,
            plugin,
            action_names: OnceLock::new(),
        }
    }

    /// `enum_action(0..)`을 끝까지 열거해 C 문자열로 변환한 목록.
    ///
    /// 이름 중간에 NUL이 있으면 C 문자열로 표현할 수 없으므로 로그를 남기고
    /// 건너뛴다.
    fn action_names(&self) -> &[CString] {
        self.action_names.get_or_init(|| {
            (0..)
                .map_while(|n| self.plugin.enum_action(n))
                .filter_map(|name| match CString::new(name) {
                    Ok(c) => Some(c),
                    Err(_) => {
                        tracing::error!(?name, "액션 이름에 NUL 문자가 있어 건너뜀");
                        None
                    }
                })
                .collect()
        })
    }
}

// 'static vtable 포인터를 보관하며 변경하지 않으므로 안전합니다.
unsafe impl<T: HwpUserAction> Sync for RustActionModule<T> {}
unsafe impl<T: HwpUserAction> Send for RustActionModule<T> {}

// =========================================================================
// 내부 헬퍼
// =========================================================================

/// C++ `this` 포인터에서 Rust 플러그인 참조를 추출합니다.
///
/// HWP가 vtable 함수를 호출할 때 넘기는 `this`는 사실
/// `RustActionModule<T>`의 첫 번째 필드(`lpVtbl`)를 가리키는 포인터입니다.
/// 해당 구조체를 재해석(`as *const RustActionModule<T>`)하면
/// 두 번째 필드(`plugin`)에 직접 접근할 수 있습니다.
///
/// # Safety
/// - `this`는 유효한 `RustActionModule<T>`를 가리켜야 합니다.
/// - 반환된 참조의 수명(`'a`)은 호출자가 책임집니다.
fn extract_plugin<'a, T: HwpUserAction>(this: *mut c_void) -> Option<&'a T> {
    extract_module::<T>(this).map(|m| &m.plugin)
}

/// C++ `this` 포인터를 `RustActionModule<T>` 참조로 재해석합니다.
///
/// # Safety
/// [`extract_plugin`]과 같습니다.
fn extract_module<'a, T: HwpUserAction>(this: *mut c_void) -> Option<&'a RustActionModule<T>> {
    if this.is_null() {
        return None;
    }
    Some(unsafe { &*(this as *const RustActionModule<T>) })
}

const MAX_ACTION_LEN: usize = 256;

fn lpcstr_to_string(ptr: *const c_char) -> String {
    if ptr.is_null() {
        return String::new();
    }
    // MAX_ACTION_LEN개 원소(u8) 내에서만 스캔해 버퍼 초과 읽기를 방지한다.
    let bytes = unsafe { std::slice::from_raw_parts(ptr as *const u8, MAX_ACTION_LEN) };
    match bytes.iter().position(|&b| b == 0) {
        Some(len) => String::from_utf8_lossy(&bytes[..len]).into_owned(),
        None => String::from_utf8_lossy(bytes).into_owned(),
    }
}

/// IDispatch를 빌려서 HwpObject를 생성하고, 작업 후 forget합니다.
/// HWP가 소유한 포인터이므로 Release하면 안 됩니다.
fn with_hwp_object<T>(p_obj: *mut c_void, f: impl FnOnce(&HwpObject) -> T) -> Option<T> {
    if p_obj.is_null() {
        return None;
    }
    let hwp = unsafe { HwpObject::from_raw_dispatch(p_obj) }.ok()?;
    let result = f(&hwp);
    std::mem::forget(hwp);
    Some(result)
}

// =========================================================================
// 트램펄린 함수 (C++ vtable → Rust trait 위임)
// =========================================================================

/// `IHncUserActionModule::EnumAction` 트램펄린
///
/// HWP가 DLL을 로드한 직후 이 함수를 `n=0`부터 반복 호출해
/// 플러그인이 제공하는 액션 이름을 순서대로 수집한다.
/// `null`을 반환하는 순간 열거가 끝난 것으로 간주한다.
///
/// # 인자 (C++ 원형: `LPCSTR EnumAction(int)`)
/// - `this` : `RustActionModule<T>` 첫 필드 주소 (ECX로 전달)
/// - `n`    : 0부터 시작하는 액션 인덱스
///
/// # 반환값
/// - null 종료 ANSI 문자열 포인터 : 해당 인덱스의 액션 이름
/// - `null ptr`                    : 열거 종료 (인덱스 범위 초과)
///
/// # 인덱스 매핑 (`enum_action` 구현 기준)
/// - `0` → `UUIDSTR_ON_INITIAL_LOAD` (예약 GUID)
/// - `1` → `UUIDSTR_ON_LOAD`         (예약 GUID)
/// - `n≥2` → `actions()[n-2].name`   (사용자 정의 액션)
///
/// # 문자열 수명
/// SDK는 반환 문자열을 언제까지 유지해야 하는지 명시하지 않는다. 그래서 첫 호출 때
/// 모든 이름을 `RustActionModule::action_names`(`static` 모듈 안의 캐시)에
/// `CString`으로 만들어 두고, 이후 그 주소를 돌려준다. 액션마다 주소가 다르고
/// DLL이 로드된 동안 내용이 바뀌지 않으므로 HWP가 포인터를 복사하든 보관하든
/// 안전하다. 길이 제한도 없다.
///
/// # Safety
/// - `this`는 `RustActionModule<T>`의 유효한 포인터이거나 `null`이어야 한다.
/// - HWP vtable 콜백으로만 호출되어야 하며, 직접 호출해서는 안 된다.
pub unsafe extern "thiscall" fn tramp_enum_action<T: HwpUserAction>(
    this: *mut c_void,
    n: i32,
) -> *const c_char {
    tracing::trace!(n, ?this, "EnumAction");
    // this가 null이면 RustActionModule에 접근할 수 없으므로 조기 반환
    let Some(module) = extract_module::<T>(this) else {
        tracing::error!("EnumAction: extract_module 실패");
        return std::ptr::null();
    };
    // n이 음수거나 범위를 벗어나면 null 반환으로 열거 종료 신호
    let names = module.action_names();
    let Some(name) = usize::try_from(n).ok().and_then(|i| names.get(i)) else {
        return std::ptr::null();
    };
    name.as_ptr()
}

/// `IHncUserActionModule::GetActionImage` 트램펄린
///
/// HWP가 툴바/리본 버튼을 렌더링할 때 각 액션의 아이콘을 요청하는 함수다.
/// 플러그인은 비트맵 스트립(960×16 BMP)의 핸들과 그 안에서의 아이콘 인덱스를 돌려준다.
///
/// # 인자 (C++ 원형: `BOOL GetActionImage(LPCSTR, UINT, HBITMAP*, int*)`)
/// - `this`          : `RustActionModule<T>` 첫 필드 주소 (ECX로 전달)
/// - `szAction`      : 아이콘을 요청하는 액션 이름 (ANSI C 문자열)
/// - `uState`        : 버튼 상태 (누름/활성/비활성 등, HWP 정의 비트마스크)
/// - `phBitmap`      : [out] 비트맵 핸들(`HBITMAP`)을 기록할 포인터
/// - `pnImageIndex`  : [out] 비트맵 스트립 내 아이콘 인덱스를 기록할 포인터
///
/// # 반환값
/// - `1` (TRUE)  : 아이콘 제공 성공 — HWP가 `*phBitmap`/`*pnImageIndex`를 사용한다
/// - `0` (FALSE) : 아이콘 없음 또는 실패 — HWP가 기본 아이콘을 표시하거나 무시한다
///
/// # 비트맵 스트립 구조
/// `toolbar_bitmap()`이 반환하는 `HBITMAP`은 960×16 픽셀 이미지다.
/// 각 아이콘은 16×16이며 `image_index * 16` 오프셋에 위치한다.
/// `pnImageIndex`에 해당 인덱스를 쓰면 HWP가 올바른 아이콘을 잘라낸다.
///
/// # Safety
/// - `this`는 `RustActionModule<T>`의 유효한 포인터이거나 `null`이어야 한다.
/// - `szAction`은 유효한 null 종료 ANSI 문자열 포인터이어야 한다.
/// - `phBitmap`, `pnImageIndex`는 유효한 포인터이거나 `null`이어야 한다.
/// - HWP vtable 콜백으로만 호출되어야 하며, 직접 호출해서는 안 된다.
#[allow(non_snake_case)]
pub unsafe extern "thiscall" fn tramp_get_action_image<T: HwpUserAction>(
    this: *mut c_void,
    szAction: *const c_char,
    uState: u32,
    phBitmap: *mut isize,
    pnImageIndex: *mut i32,
) -> i32 {
    // this가 null이면 RustActionModule에 접근할 수 없으므로 조기 반환
    let Some(plugin) = extract_plugin::<T>(this) else {
        tracing::error!("GetActionImage: extract_plugin 실패");
        return 0;
    };
    let action_name = lpcstr_to_string(szAction);
    tracing::trace!(action = %action_name, state = uState, "GetActionImage");

    // get_action_image: actions() 목록에서 액션을 찾고, toolbar_bitmap()으로 HBITMAP을 얻음.
    // 해당 액션이 없거나 비트맵 미설정이면 None → FALSE 반환
    if let Some((hbitmap, index)) = plugin.get_action_image(&action_name, uState) {
        tracing::trace!(hbitmap, index, "GetActionImage");
        // 출력 포인터가 null인 경우도 허용 (HWP가 특정 인자만 사용할 수 있음)
        if !phBitmap.is_null() {
            unsafe { *phBitmap = hbitmap };
        }
        if !pnImageIndex.is_null() {
            unsafe { *pnImageIndex = index };
        }
        1
    } else {
        tracing::trace!("GetActionImage: None 반환");
        0
    }
}

/// `IHncUserActionModule::UpdateUI` 트램펄린
///
/// HWP가 버튼/메뉴 상태를 갱신할 때 각 액션마다 이 함수를 호출한다.
/// 반환값과 `*lpuState`에 따라 버튼이 활성/비활성/토글 상태로 표시된다.
///
/// # 인자 (C++ 원형: `BOOL UpdateUI(LPCSTR, LPDISPATCH, UINT*)`)
/// - `this`      : `RustActionModule<T>` 첫 필드 주소 (ECX로 전달)
/// - `szAction`  : 상태를 조회할 액션 이름 (ANSI C 문자열)
/// - `pObj`      : HWP `IDispatch` 포인터 (HWP 소유, Release 금지)
/// - `lpuState`  : [out] 버튼 UI 상태 비트마스크를 기록할 포인터
///
/// # 반환값
/// - `1` (TRUE)  : 상태 조회 성공 — HWP가 `*lpuState` 값을 반영한다
/// - `0` (FALSE) : 실패 (null this 또는 null pObj) — HWP가 상태를 무시한다
///
/// # Safety
/// - `this`는 `RustActionModule<T>`의 유효한 포인터이거나 `null`이어야 한다.
/// - `szAction`은 유효한 null 종료 ANSI 문자열 포인터이어야 한다.
/// - `pObj`는 유효한 `IDispatch` 포인터이거나 `null`이어야 한다 (Release 금지).
/// - `lpuState`는 유효한 포인터이거나 `null`이어야 한다.
/// - HWP vtable 콜백으로만 호출되어야 하며, 직접 호출해서는 안 된다.
#[allow(non_snake_case)]
pub unsafe extern "thiscall" fn tramp_update_ui<T: HwpUserAction>(
    this: *mut c_void,
    szAction: *const c_char,
    pObj: *mut c_void,
    lpuState: *mut u32,
) -> i32 {
    // this가 null이면 RustActionModule에 접근할 수 없으므로 조기 반환
    let Some(plugin) = extract_plugin::<T>(this) else {
        return 0;
    };
    let action_name = lpcstr_to_string(szAction);

    // pObj(IDispatch)를 HwpObject로 래핑해 트레잇 메서드에 전달.
    // pObj가 null이면 with_hwp_object가 None을 반환 → FALSE 반환.
    let Some(_) = (unsafe {
        with_hwp_object(pObj, |hwp| {
            // dispatch_update_ui: 예약 GUID → 0, 사용자 액션 → update_ui() 위임
            let state = plugin.dispatch_update_ui(&action_name, hwp);
            // lpuState가 null이 아닐 때만 출력 인자에 기록
            if !lpuState.is_null() {
                *lpuState = state;
            }
        })
    }) else {
        return 0;
    };
    1
}

/// `IHncUserActionModule::DoAction` 트램펄린
///
/// 사용자가 버튼을 클릭하거나 단축키를 누르는 등 액션이 실행될 때
/// HWP가 이 함수를 호출한다. 플러그인의 핵심 진입점이다.
///
/// # 인자 (C++ 원형: `int DoAction(LPCSTR, LPDISPATCH)`)
/// - `this`     : `RustActionModule<T>` 첫 필드 주소 (ECX로 전달)
/// - `szAction` : 실행할 액션 이름 (ANSI C 문자열)
/// - `pObj`     : HWP `IDispatch` 포인터 (HWP 소유, Release 금지)
///
/// # 반환값
/// - `1` (성공)  : 액션이 정상 처리됨 — `dispatch`가 `Ok(true)`를 반환한 경우
/// - `0` (실패)  : 다음 중 하나:
///     - `this`가 null (extract_plugin 실패)
///     - `pObj`가 null (HWP 객체 래핑 불가)
///     - `dispatch`가 `Err(_)` 또는 `Ok(false)`를 반환
///
/// # dispatch 라우팅
/// `plugin.dispatch()`는 액션 이름에 따라 다음으로 위임한다:
/// - `UUIDSTR_ON_INITIAL_LOAD` → `on_initial_load()` (최초 등록 시 1회)
/// - `UUIDSTR_ON_LOAD`         → `on_load()` → `setup_toolbar()` (창 열릴 때마다)
/// - 그 외                     → `do_action()` (사용자 정의 액션)
///
/// # Safety
/// - `this`는 `RustActionModule<T>`의 유효한 포인터이거나 `null`이어야 한다.
/// - `szAction`은 유효한 null 종료 ANSI 문자열 포인터이어야 한다.
/// - `pObj`는 유효한 `IDispatch` 포인터이거나 `null`이어야 한다 (Release 금지).
/// - HWP vtable 콜백으로만 호출되어야 하며, 직접 호출해서는 안 된다.
#[allow(non_snake_case)]
pub unsafe extern "thiscall" fn tramp_do_action<T: HwpUserAction>(
    this: *mut c_void,
    szAction: *const c_char,
    pObj: *mut c_void,
) -> i32 {
    // this가 null이면 RustActionModule에 접근할 수 없으므로 조기 반환
    let Some(plugin) = extract_plugin::<T>(this) else {
        return 0;
    };
    let action_name = lpcstr_to_string(szAction);
    // 이 액션 안에서 나온 로그(hwp_core의 COM 호출 포함)에 action 이름을 붙인다.
    let _span = tracing::debug_span!("DoAction", action = %action_name).entered();
    tracing::debug!("DoAction");

    // pObj(IDispatch)를 HwpObject로 빌려서 dispatch에 전달.
    // with_hwp_object 내부에서 mem::forget으로 Drop(=Release)을 방지한다.
    let result = with_hwp_object(pObj, |hwp| plugin.dispatch(&action_name, hwp));

    // 에러/실패 경로는 모두 로그에 기록한다.
    // Some(Ok(true))만 성공으로 간주하고 나머지는 0을 반환한다.
    match &result {
        Some(Ok(true)) => {}
        Some(Ok(false)) => tracing::debug!("Ok(false) 반환"),
        Some(Err(e)) => tracing::error!(error = %e, "DoAction 실패"),
        None => tracing::error!("with_hwp_object None (p_obj null 또는 래핑 실패)"),
    }

    match result {
        Some(Ok(true)) => 1,
        _ => 0,
    }
}
