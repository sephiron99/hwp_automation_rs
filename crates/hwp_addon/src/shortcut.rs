//! 단축키 제공 — Windows Keyboard Hook (HWP 스레드 한정)
//!
//! 키보드 훅은 단축키 감지만 담당하고, 매칭된 액션은 **PostMessage**로
//! 전용 메시지 윈도우에 신호를 보낸다. 실제 `do_action` 실행은 메시지 윈도우의
//! `wndproc`에서 일어나므로, 훅 콜백 컨텍스트(HWP의 키 처리 도중)를 벗어난
//! 평범한 메시지 디스패치 흐름에서 실행된다. 이러면 새 윈도우 활성화·foreground
//! 전환 등이 정상 동작한다.

use std::ffi::c_void;
use std::sync::Mutex;
use std::sync::atomic::{AtomicBool, AtomicPtr, Ordering};

use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::System::Threading::GetCurrentThreadId;
use windows::Win32::UI::Input::KeyboardAndMouse::{
    GetAsyncKeyState, VIRTUAL_KEY, VK_CONTROL, VK_MENU, VK_SHIFT,
};
use windows::Win32::UI::WindowsAndMessaging::{
    CallNextHookEx, CreateWindowExW, DefWindowProcW, HHOOK, HWND_MESSAGE, PostMessageW,
    RegisterClassW, SetWindowsHookExW, WH_KEYBOARD, WINDOW_EX_STYLE, WINDOW_STYLE, WM_USER,
    WNDCLASSW,
};
use windows::core::w;

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

/// PostMessage 대상 message-only 윈도우.
static MESSAGE_HWND: AtomicPtr<c_void> = AtomicPtr::new(std::ptr::null_mut());

/// addon 단축키 액션 dispatch용 user 메시지.
const WM_ADDON_ACTION: u32 = WM_USER + 0x1234;

/// addon 측이 자체 메시지 펌프를 돌릴 때 hook이 단축키를 swallow하지 않도록
/// 켜는 플래그. true이면 hook은 모든 키를 그대로 통과시킨다.
static POPUP_ACTIVE: AtomicBool = AtomicBool::new(false);

/// popup·dialog 등 자체 메시지 펌프를 돌리는 동안 호출해 hook을 우회시킨다.
pub fn set_popup_active(active: bool) {
    POPUP_ACTIVE.store(active, Ordering::Relaxed);
}

fn is_popup_active() -> bool {
    POPUP_ACTIVE.load(Ordering::Relaxed)
}

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
/// 메시지 윈도우의 wndproc에서 호출됩니다. 즉 훅 콜백 컨텍스트를 벗어난
/// 평범한 메시지 디스패치 시점이므로 새 윈도우 활성화·foreground 전환 등
/// 일반 UI 동작이 정상적으로 가능합니다.
///
/// 단축키 경로는 `HwpUserAction::dispatch()`를 거치지 않고 `plugin.do_action`을
/// 직접 호출하므로, IME 조합 확정(commit)도 여기서 수행해야 합니다.
fn run_action(action_name: &str) -> Result<bool> {
    let raw = HWP_DISPATCH.load(Ordering::Relaxed);
    if raw.is_null() {
        return Ok(false);
    }
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

// =========================================================================
// 메시지 윈도우 (PostMessage 타깃)
// =========================================================================

unsafe extern "system" fn message_wnd_proc(
    hwnd: HWND,
    msg: u32,
    wp: WPARAM,
    lp: LPARAM,
) -> LRESULT {
    if msg == WM_ADDON_ACTION {
        let index = wp.0;
        let name_opt: Option<&'static str> = ACTIONSWITHSHORTCUT
            .lock()
            .ok()
            .and_then(|v| v.get(index).map(|m| m.name));
        if let Some(name) = name_opt {
            let _ = run_action(name);
        }
        return LRESULT(0);
    }
    unsafe { DefWindowProcW(hwnd, msg, wp, lp) }
}

unsafe fn create_message_window() -> HWND {
    unsafe {
        let hmodule = GetModuleHandleW(None).unwrap_or_default();
        let hinstance = HINSTANCE(hmodule.0);
        let class_name = w!("HwpAddonMessageWindow");

        let wc = WNDCLASSW {
            hInstance: hinstance,
            lpszClassName: class_name,
            lpfnWndProc: Some(message_wnd_proc),
            ..Default::default()
        };
        RegisterClassW(&wc);

        let hwnd = CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            class_name,
            w!(""),
            WINDOW_STYLE::default(),
            0,
            0,
            0,
            0,
            Some(HWND_MESSAGE),
            None,
            Some(hinstance),
            None,
        );
        hwnd.unwrap_or(HWND(std::ptr::null_mut()))
    }
}

// =========================================================================
// Hook 설치 및 콜백
// =========================================================================

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
        unsafe {
            // 1) PostMessage 타깃 메시지 윈도우 생성
            let mw = create_message_window();
            MESSAGE_HWND.store(mw.0, Ordering::Relaxed);

            // 2) HWP 스레드에 키보드 훅 설치
            let tid = GetCurrentThreadId();
            let _hook = SetWindowsHookExW(WH_KEYBOARD, Some(keyboard_hook_proc), None, tid)
                .expect("키보드 훅 설치 실패");
        }
    }
}

/// WH_KEYBOARD 훅 콜백 함수.
///
/// `wParam`은 가상 키 코드, `lParam` bit 31은 전환 상태(0=누름, 1=뗌)입니다.
/// 단축키가 매칭되면 PostMessage로 메시지 윈도우에 신호만 보내고 1을 반환해
/// keystroke를 소비합니다. 실제 `do_action`은 메시지 윈도우의 wndproc에서
/// 비동기적으로 실행됩니다.
///
/// `set_popup_active(true)`가 호출된 동안에는 hook은 동작하지 않고 모든 키를
/// 그대로 다음 hook으로 흘려보냅니다. addon 측 popup 메시지 펌프가 자체적으로
/// 키 처리를 하도록 양보하기 위함입니다.
extern "system" fn keyboard_hook_proc(n_code: i32, w_param: WPARAM, l_param: LPARAM) -> LRESULT {
    if is_popup_active() {
        return unsafe { CallNextHookEx(Some(HHOOK::default()), n_code, w_param, l_param) };
    }

    // bit 31 == 0 → 키 누름
    let is_key_down = (l_param.0 as u32 >> 31) == 0;
    if n_code >= 0 && is_key_down {
        let vk = VIRTUAL_KEY(w_param.0 as u16);
        let current_mods = Modifiers::current();

        let matched_index = if let Ok(shortcuts) = ACTIONSWITHSHORTCUT.lock() {
            shortcuts.iter().position(|sc| {
                sc.shortcut
                    .map(|s| s.key == vk && s.modifiers == current_mods)
                    .unwrap_or(false)
            })
        } else {
            None
        };

        if let Some(index) = matched_index {
            let mw = MESSAGE_HWND.load(Ordering::Relaxed);
            if !mw.is_null() {
                unsafe {
                    let _ = PostMessageW(
                        Some(HWND(mw)),
                        WM_ADDON_ACTION,
                        WPARAM(index),
                        LPARAM(0),
                    );
                }
            }
            return LRESULT(1); // keystroke 소비
        }
    }

    unsafe { CallNextHookEx(Some(HHOOK::default()), n_code, w_param, l_param) }
}
