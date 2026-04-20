//! 단축키 제공 — Windows Keyboard Hook (HWP 스레드 한정)

use std::ffi::c_void;
use std::sync::Mutex;
use std::sync::atomic::{AtomicPtr, Ordering};
use windows::Win32::Foundation::{LPARAM, LRESULT, WPARAM};
use windows::Win32::System::Threading::GetCurrentThreadId;
use windows::Win32::UI::Input::KeyboardAndMouse::{
    GetAsyncKeyState, VIRTUAL_KEY, VK_CONTROL, VK_MENU, VK_SHIFT,
};
use windows::Win32::UI::WindowsAndMessaging::{
    CallNextHookEx, HHOOK, SetWindowsHookExW, WH_KEYBOARD,
};

use hwp_core::error::Result;
use hwp_core::hwp_obj::HwpObject;

use crate::hwp_user_action::ActionMeta;
// HwpUserAction는 set_action_callback<T: HwpUserAction>에서 사용
use crate::hwp_user_action::HwpUserAction;

/// 수식 키 조합.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct Modifiers {
    pub alt: bool,
    pub ctrl: bool,
    pub shift: bool,
}

impl Modifiers {
    /// 현재 키보드에서 눌려있는 수식 키 상태를 조회합니다.
    fn current() -> Self {
        unsafe {
            Self {
                alt: GetAsyncKeyState(VK_MENU.0.into()).is_negative(),
                ctrl: GetAsyncKeyState(VK_CONTROL.0.into()).is_negative(),
                shift: GetAsyncKeyState(VK_SHIFT.0.into()).is_negative(),
            }
        }
    }
}

/// `ActionMeta`에 지정할 단축키 조합.
///
/// # Example
/// ```ignore
/// ActionMeta {
///     name: "MyAction",
///     label: "내 액션",
///     image_index: 0,
///     shortcut: Some(ShortcutKey {
///         modifiers: Modifiers { alt: true, ctrl: false, shift: false },
///         key: VK_D,
///     }),
/// }
/// ```
#[derive(Clone, Copy, Debug)]
pub struct ShortcutKey {
    pub modifiers: Modifiers,
    pub key: VIRTUAL_KEY,
}

/// 단축키가 등록된 목록.
static ACTIONSWITHSHORTCUT: Mutex<Vec<ActionMeta>> = Mutex::new(Vec::new());

// =========================================================================
// HWP IDispatch 저장 + 액션 실행
// =========================================================================

/// 현재 HWP IDispatch raw 포인터. `on_load` 시 갱신됩니다.
static HWP_DISPATCH: AtomicPtr<c_void> = AtomicPtr::new(std::ptr::null_mut());

/// HWP IDispatch 포인터를 저장합니다.
pub(crate) fn set_hwp_dispatch(hwp: &HwpObject) {
    HWP_DISPATCH.store(hwp.as_raw_dispatch(), Ordering::Relaxed);
}

// =========================================================================
// 플러그인 콜백 — do_action 직접 호출
// =========================================================================

/// 플러그인 `do_action`을 호출하기 위한 type-erased 콜백.
///
/// `plugin_ptr`는 `static RustActionModule` 내 플러그인을 가리키므로
/// 프로세스 종료까지 유효합니다.
struct ActionCallback {
    plugin_ptr: *const c_void,
    call_fn: fn(*const c_void, &str, &HwpObject) -> Result<bool>,
}

// SAFETY: plugin_ptr은 static에 저장된 HwpUserAction을 가리킵니다.
unsafe impl Send for ActionCallback {}
unsafe impl Sync for ActionCallback {}

static ACTION_CALLBACK: Mutex<Option<ActionCallback>> = Mutex::new(None);

/// 플러그인의 `do_action`을 단축키에서 호출할 수 있도록 등록합니다.
///
/// 제네릭 `T`를 단형화(monomorphize)하여 type-erased 함수 포인터로 저장합니다.
pub(crate) fn set_action_callback<T: HwpUserAction>(plugin: &T) {
    fn call<T: HwpUserAction>(ptr: *const c_void, name: &str, hwp: &HwpObject) -> Result<bool> {
        let plugin = unsafe { &*(ptr as *const T) };
        plugin.do_action(name, hwp)
    }

    *ACTION_CALLBACK.lock().unwrap() = Some(ActionCallback {
        plugin_ptr: plugin as *const T as *const c_void,
        call_fn: call::<T>,
    });
}

/// 저장된 콜백으로 플러그인 액션을 실행합니다.
///
/// 단축키 경로는 `HwpUserAction::dispatch()`를 거치지 않고 `plugin.do_action`을
/// 직접 호출하므로, IME 조합 확정(commit)도 여기서 수행해야 합니다. 한글 조합
/// 중 단축키 입력 시 후속 `Move*`/`Select*` 액션이 캐럿을 움직이지 못하는
/// 버그 방지용입니다.
fn run_action(action_name: &str) -> Result<bool> {
    let raw = HWP_DISPATCH.load(Ordering::Relaxed);
    if raw.is_null() {
        return Ok(false);
    }
    // 키보드 훅 콜백 내부에서 호출되므로 여기가 `do_action` 진입 전에 IME
    // 상태에 개입할 수 있는 유일한 지점입니다.
    crate::ime::commit_composition();
    let hwp = unsafe { HwpObject::from_raw_dispatch(raw) }?;
    let result = if let Ok(cbguard) = ACTION_CALLBACK.lock()
        && let Some(ref cb) = *cbguard
    {
        (cb.call_fn)(cb.plugin_ptr, action_name, &hwp)?
    } else {
        false
    };
    std::mem::forget(hwp);
    Ok(result)
}

/// `ActionMeta` 목록에서 단축키가 지정된 항목을 등록하고, 키보드 훅을 설치합니다.
///
/// 동일 스레드(HWP 스레드)에 훅을 설치하므로 훅 콜백·타이머 콜백 모두
/// HWP 스레드에서 실행되며, `SetKeyboardState`로 HWP 스레드의 키 상태를
/// 직접 변경할 수 있습니다.
///
/// 최초 1회만 실행됩니다.
pub(crate) fn register_action_shortcuts(actions: &impl HwpUserAction) {
    // 단축키가 지정된 액션만 ACTIONSWITHSHORTCUT에 모음
    for shc in actions
        .actions()
        .iter()
        .filter(|act| act.shortcut.is_some())
    {
        ACTIONSWITHSHORTCUT.lock().unwrap().push(shc.clone());
    }

    if !ACTIONSWITHSHORTCUT.lock().unwrap().is_empty() {
        // 동일 스레드(HWP 스레드)에 훅을 설치합니다.
        // HWP의 메시지 루프가 이미 존재하므로 별도 메시지 루프 불필요.
        unsafe {
            let tid = GetCurrentThreadId();
            let _hook = SetWindowsHookExW(WH_KEYBOARD, Some(keyboard_hook_proc), None, tid)
                .expect("키보드 훅 설치 실패");
            // *HOOK_HANDLE.lock().unwrap() = Some(HookHandle(hook));
        }
    }
}

/// WH_KEYBOARD 훅 콜백 함수.
///
/// `wParam`은 가상 키 코드, `lParam` bit 31은 전환 상태(0=누름, 1=뗌)입니다.
/// 단축키 일치 시 keystroke를 소비하고 `SetTimer(0)`으로 지연 실행을 예약합니다.
///
/// 지연 실행이 필요한 이유: hook 콜백은 HWP의 키보드 메시지 처리 도중에 호출되므로,
/// 이 안에서 HWP COM 액션을 직접 실행하면 HWP 내부 상태가 불안정할 수 있습니다.
/// `SetTimer(0)`으로 현재 메시지 처리가 완료된 후 실행하면 이 문제를 회피합니다.
extern "system" fn keyboard_hook_proc(n_code: i32, w_param: WPARAM, l_param: LPARAM) -> LRESULT {
    // bit 31 == 0 → 키 누름
    let is_key_down = (l_param.0 as u32 >> 31) == 0;
    if n_code >= 0 && is_key_down {
        let vk = VIRTUAL_KEY(w_param.0 as u16);
        let current_mods = Modifiers::current();

        if let Ok(shortcuts) = ACTIONSWITHSHORTCUT.lock()
            && let Some(shortcut) = shortcuts.iter().find(|sc| {
                sc.shortcut
                    .map(|s| s.key == vk && s.modifiers == current_mods)
                    .unwrap_or(false)
            })
        {
            // extern "system" 제약상 let _ = run_action(...) 으로 무시
            let _ = run_action(shortcut.name);
            // keystroke를 소비하여 HWP가 중복 처리하지 않도록 합니다.
            return LRESULT(1);
        }
    }

    // 매칭 안 되면 HWP 기본 처리
    unsafe { CallNextHookEx(Some(HHOOK::default()), n_code, w_param, l_param) }
}
