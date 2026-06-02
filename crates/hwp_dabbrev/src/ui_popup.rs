//! Dabbrev 자동완성 popup (Win32 전용).
//!
//! Modal 윈도우로 candidate 목록을 보여주고, 키별로 cycle/확정/취소 동작을
//! 분기한다. message pump는 `show()` 내부에서 직접 운영하며, 키 처리는
//! listbox로 dispatch하기 전에 가로채서 한다.
//!
//! ## 키 매핑
//! - ↑ / ↓ / Ctrl+/ — preview replace + popup 유지 (cycle)
//! - Enter — 확정, popup 닫음, 키 swallow
//! - Right / Space — 확정, popup 닫음, 키를 HWP로 forward
//! - ESC — 취소, popup 닫음, 키 swallow
//! - 그 외 모든 키 — 취소, popup 닫음, 키를 HWP로 forward

use std::cell::{Cell, RefCell};
use std::ffi::c_void;

use windows::Win32::Foundation::{HINSTANCE, HWND, LPARAM, LRESULT, POINT, WPARAM};
use windows::Win32::Graphics::Gdi::{COLOR_WINDOW, ClientToScreen, HBRUSH, UpdateWindow};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Input::KeyboardAndMouse::{
    GetKeyState, GetKeyboardLayout, GetKeyboardState, SetActiveWindow, SetFocus, ToUnicodeEx,
    VIRTUAL_KEY, VK_CONTROL, VK_DOWN, VK_ESCAPE, VK_OEM_2, VK_RETURN, VK_RIGHT, VK_SPACE, VK_UP,
};
use windows::Win32::UI::WindowsAndMessaging::{
    BringWindowToTop, CreateWindowExW, DefWindowProcW, DestroyWindow, DispatchMessageW,
    GUITHREADINFO, GetForegroundWindow, GetGUIThreadInfo, GetMessageW, GetSystemMetrics,
    GetWindowThreadProcessId, HMENU, IDC_ARROW, LB_ADDSTRING, LB_GETCOUNT, LB_GETCURSEL,
    LB_SETCURSEL, LBS_NOTIFY, LoadCursorW, MSG, PostMessageW, RegisterClassW, SM_CXSCREEN,
    SM_CYSCREEN, SW_SHOWNORMAL, SendMessageW, SetForegroundWindow, ShowWindow, WINDOW_EX_STYLE,
    WINDOW_STYLE, WM_ACTIVATE, WM_CHAR, WM_DESTROY, WM_KEYDOWN, WM_KEYUP, WNDCLASSW, WS_BORDER,
    WS_CHILD, WS_EX_TOOLWINDOW, WS_EX_TOPMOST, WS_POPUP, WS_VISIBLE, WS_VSCROLL,
};
use windows::core::w;

const POPUP_WIDTH: i32 = 200;
const POPUP_HEIGHT: i32 = 120;

const ID_LIST: isize = 201;

/// popup 종료 원인.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Outcome {
    /// Enter / Right / Space — 명시적 확정.
    Committed,
    /// ESC / 그 외 키 / 닫힘 — 취소.
    Cancelled,
}

type FetchMore = Box<dyn FnMut() -> Vec<String>>;
type ReplaceCb = Box<dyn FnMut(usize)>;

thread_local! {
    static OUTCOME: Cell<Outcome> = const { Cell::new(Outcome::Cancelled) };
    static LIST_HWND: Cell<isize> = const { Cell::new(0) };
    static FETCH_MORE: RefCell<Option<FetchMore>> = const { RefCell::new(None) };
    static REPLACE_CB: RefCell<Option<ReplaceCb>> = const { RefCell::new(None) };
    static EXHAUSTED: Cell<bool> = const { Cell::new(false) };
    static INITIAL: RefCell<Vec<String>> = const { RefCell::new(Vec::new()) };
    static START_INDEX: Cell<usize> = const { Cell::new(0) };
    /// WM_DESTROY 도착 → nested pump가 빠져나오도록 set. PostQuitMessage 대신
    /// 사용하는 이유는 PostQuitMessage가 HWP main pump까지 종료시키기 때문.
    static POPUP_DONE: Cell<bool> = const { Cell::new(false) };
    /// 처음 WM_ACTIVATE(WA_ACTIVE)를 받은 후에만 deactivate를 close 트리거로 본다.
    /// 초기 생성 시점의 transient deactivate에 즉시 닫히지 않도록.
    static GOT_ACTIVE: Cell<bool> = const { Cell::new(false) };
    /// 키 분기에서 DestroyWindow를 호출하기 직전에 set. DestroyWindow가 동기적으로
    /// 발생시키는 WM_ACTIVATE(WA_INACTIVE)가 outcome을 덮어쓰지 못하도록 가드.
    static EXPLICIT_CLOSE: Cell<bool> = const { Cell::new(false) };
}

/// popup을 띄우고 사용자 조작이 끝날 때까지 블록한다.
///
/// - `initial` — 첫 batch candidates. `start_index`는 첫 sel 위치 (preview replace는
///   이미 호출자가 수행했음).
/// - `fetch_more` — listbox 끝에서 사용자가 더 내려갈 때 호출. 빈 Vec 반환 시
///   더 호출하지 않음.
/// - `replace` — sel이 바뀔 때마다 호출. 호출자가 문서에 preview replace 수행.
/// - `forward_target` — 확정/취소 후 키를 forward할 윈도우 (HWP 메인).
pub fn show(
    initial: &[String],
    start_index: usize,
    fetch_more: impl FnMut() -> Vec<String> + 'static,
    replace: impl FnMut(usize) + 'static,
    forward_target: HWND,
) -> Outcome {
    OUTCOME.with(|o| o.set(Outcome::Cancelled));
    EXHAUSTED.with(|e| e.set(false));
    GOT_ACTIVE.with(|g| g.set(false));
    EXPLICIT_CLOSE.with(|e| e.set(false));
    INITIAL.with(|i| *i.borrow_mut() = initial.to_vec());
    START_INDEX.with(|s| s.set(start_index));
    FETCH_MORE.with(|f| *f.borrow_mut() = Some(Box::new(fetch_more)));
    REPLACE_CB.with(|r| *r.borrow_mut() = Some(Box::new(replace)));

    // hook이 단축키를 가로채지 않도록 양보 — popup 펌프가 직접 Ctrl+/ 처리.
    hwp_addon::shortcut::set_popup_active(true);

    let popup = unsafe { create_popup(forward_target) };
    if popup.0.is_null() {
        hwp_addon::shortcut::set_popup_active(false);
        cleanup();
        return Outcome::Cancelled;
    }

    unsafe {
        let _ = ShowWindow(popup, SW_SHOWNORMAL);
        let _ = UpdateWindow(popup);
        let _ = BringWindowToTop(popup);
        let _ = SetForegroundWindow(popup);
        let _ = SetActiveWindow(popup);
        let list = HWND(LIST_HWND.with(|l| l.get()) as *mut c_void);
        let _ = SetFocus(Some(list));

        run_pump(popup, forward_target);

        // popup destroy 후 HWP를 다시 활성화. ShowWindow 계열은 절대 호출하지
        // 않는다 — SW_RESTORE는 maximize 상태를 normal로 되돌리면서 위치를
        // 옮기고, SW_SHOWNORMAL도 동일한 부작용이 있다. BringWindowToTop과
        // SetForegroundWindow는 z-order/activation만 바꾸므로 안전.
        if !forward_target.0.is_null() {
            let _ = BringWindowToTop(forward_target);
            let _ = SetForegroundWindow(forward_target);
        }
    }

    hwp_addon::shortcut::set_popup_active(false);
    let outcome = OUTCOME.with(|o| o.get());
    cleanup();
    outcome
}

fn cleanup() {
    FETCH_MORE.with(|f| *f.borrow_mut() = None);
    REPLACE_CB.with(|r| *r.borrow_mut() = None);
    INITIAL.with(|i| i.borrow_mut().clear());
    LIST_HWND.with(|l| l.set(0));
}

unsafe fn create_popup(owner: HWND) -> HWND {
    unsafe {
        let hmodule = GetModuleHandleW(None).unwrap_or_default();
        let hinstance = HINSTANCE(hmodule.0);
        let class_name = w!("HwpDabbrevPopup");

        let wc = WNDCLASSW {
            hCursor: LoadCursorW(None, IDC_ARROW).unwrap_or_default(),
            hInstance: hinstance,
            lpszClassName: class_name,
            lpfnWndProc: Some(wnd_proc),
            hbrBackground: HBRUSH((COLOR_WINDOW.0 as usize + 1) as *mut c_void),
            ..Default::default()
        };
        // ATOM이 0이어도(이미 등록됨) 다음 CreateWindowExW에서 사용 가능.
        RegisterClassW(&wc);

        let anchor = try_anchor();
        // owner를 HWP 메인 윈도우로 지정 — popup 종료 시 focus가 HWP로 자동 복귀.
        let owner_opt = if owner.0.is_null() { None } else { Some(owner) };
        let popup = CreateWindowExW(
            WS_EX_TOPMOST | WS_EX_TOOLWINDOW,
            class_name,
            w!("dabbrev"),
            WS_POPUP | WS_BORDER,
            anchor.x,
            anchor.y,
            POPUP_WIDTH,
            POPUP_HEIGHT,
            owner_opt,
            None,
            Some(hinstance),
            None,
        );
        popup.unwrap_or(HWND(std::ptr::null_mut()))
    }
}

/// 캐럿 화면 좌표 시도. 실패 시 화면 중앙.
///
/// 캐럿의 화면 y가 화면 높이의 2/3보다 더 낮은 위치(= 하단 1/3 영역)에 있으면
/// popup을 캐럿 위에 띄워 모니터 밖 이탈을 막는다. 그 외엔 캐럿 아래.
unsafe fn try_anchor() -> POINT {
    unsafe {
        let screen_w = GetSystemMetrics(SM_CXSCREEN);
        let screen_h = GetSystemMetrics(SM_CYSCREEN);
        // popup 오른쪽이 화면 밖으로 잘리지 않도록 x 최대값 제한 (왼쪽은 0).
        let max_x = (screen_w - POPUP_WIDTH).max(0);
        let clamp_x = |x: i32| x.clamp(0, max_x);
        let fg = GetForegroundWindow();
        if !fg.0.is_null() {
            let tid = GetWindowThreadProcessId(fg, None);
            if tid != 0 {
                let mut info = GUITHREADINFO {
                    cbSize: std::mem::size_of::<GUITHREADINFO>() as u32,
                    ..Default::default()
                };
                if GetGUIThreadInfo(tid, &mut info).is_ok()
                    && !info.hwndCaret.0.is_null()
                    && (info.rcCaret.right - info.rcCaret.left) >= 0
                    && (info.rcCaret.bottom - info.rcCaret.top) > 0
                {
                    let mut top_pt = POINT {
                        x: info.rcCaret.left,
                        y: info.rcCaret.top,
                    };
                    let mut bot_pt = POINT {
                        x: info.rcCaret.left,
                        y: info.rcCaret.bottom,
                    };
                    if ClientToScreen(info.hwndCaret, &mut top_pt).as_bool()
                        && ClientToScreen(info.hwndCaret, &mut bot_pt).as_bool()
                    {
                        let threshold = screen_h * 2 / 3;
                        let caret_h = (bot_pt.y - top_pt.y).max(0);
                        let margin = caret_h.max(40);
                        if bot_pt.y > threshold {
                            // 캐럿이 화면 하단 1/3에 있음 → 위로. 음수로 빠지면
                            // 화면 상단으로 clamp.
                            let y = top_pt.y - POPUP_HEIGHT - margin;
                            return POINT {
                                x: clamp_x(top_pt.x),
                                y,
                            };
                        } else {
                            let y = bot_pt.y + 2;
                            return POINT {
                                x: clamp_x(bot_pt.x),
                                y,
                            };
                        }
                    }
                }
            }
        }
        POINT {
            x: clamp_x(screen_w / 2 - POPUP_WIDTH / 2),
            y: screen_h / 2 - POPUP_HEIGHT / 2,
        }
    }
}

unsafe extern "system" fn wnd_proc(hwnd: HWND, msg: u32, wp: WPARAM, lp: LPARAM) -> LRESULT {
    unsafe {
        match msg {
            x if x == 0x0001 => {
                // WM_CREATE — listbox 생성.
                on_create(hwnd);
                LRESULT(0)
            }
            WM_ACTIVATE => {
                // wParam LOWORD: 0=WA_INACTIVE, 1=WA_ACTIVE, 2=WA_CLICKACTIVE.
                let state = (wp.0 & 0xFFFF) as u32;
                if state == 0 {
                    // EXPLICIT_CLOSE: 키 분기에서 우리가 직접 DestroyWindow한 경우
                    // — 이미 outcome이 정해졌으므로 건드리지 않는다.
                    if GOT_ACTIVE.with(|g| g.get()) && !EXPLICIT_CLOSE.with(|e| e.get()) {
                        // 외부 요인(Alt+Tab, 외부 클릭 등)으로 활성 잃음 → cancel.
                        OUTCOME.with(|o| o.set(Outcome::Cancelled));
                        let _ = DestroyWindow(hwnd);
                    }
                } else {
                    GOT_ACTIVE.with(|g| g.set(true));
                }
                LRESULT(0)
            }
            WM_DESTROY => {
                POPUP_DONE.with(|d| d.set(true));
                LRESULT(0)
            }
            _ => DefWindowProcW(hwnd, msg, wp, lp),
        }
    }
}

unsafe fn on_create(hwnd: HWND) {
    unsafe {
        let hmodule = GetModuleHandleW(None).unwrap_or_default();
        let hinstance = HINSTANCE(hmodule.0);
        let list = CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            w!("LISTBOX"),
            w!(""),
            WS_CHILD | WS_VISIBLE | WS_VSCROLL | WINDOW_STYLE(LBS_NOTIFY as u32),
            0,
            0,
            POPUP_WIDTH - 2,
            POPUP_HEIGHT - 2,
            Some(hwnd),
            Some(HMENU(ID_LIST as *mut c_void)),
            Some(hinstance),
            None,
        )
        .unwrap_or(HWND(std::ptr::null_mut()));
        LIST_HWND.with(|l| l.set(list.0 as isize));

        if !list.0.is_null() {
            INITIAL.with(|i| {
                for s in i.borrow().iter() {
                    push_listbox_item(list, s);
                }
            });
            let start = START_INDEX.with(|s| s.get()) as isize;
            SendMessageW(
                list,
                LB_SETCURSEL,
                Some(WPARAM(start as usize)),
                Some(LPARAM(0)),
            );
        }
    }
}

unsafe fn push_listbox_item(list: HWND, text: &str) {
    unsafe {
        let wide: Vec<u16> = text.encode_utf16().chain(std::iter::once(0)).collect();
        SendMessageW(
            list,
            LB_ADDSTRING,
            Some(WPARAM(0)),
            Some(LPARAM(wide.as_ptr() as isize)),
        );
    }
}

unsafe fn list_count(list: HWND) -> i32 {
    unsafe { SendMessageW(list, LB_GETCOUNT, Some(WPARAM(0)), Some(LPARAM(0))).0 as i32 }
}

unsafe fn list_cursel(list: HWND) -> i32 {
    unsafe { SendMessageW(list, LB_GETCURSEL, Some(WPARAM(0)), Some(LPARAM(0))).0 as i32 }
}

unsafe fn list_setsel(list: HWND, idx: i32) {
    unsafe {
        SendMessageW(
            list,
            LB_SETCURSEL,
            Some(WPARAM(idx as usize)),
            Some(LPARAM(0)),
        );
    }
}

fn is_ctrl_down() -> bool {
    unsafe { (GetKeyState(VK_CONTROL.0 as i32) as u16) & 0x8000 != 0 }
}

fn is_ctrl_slash(msg: &MSG) -> bool {
    let vk = msg.wParam.0 as u16;
    vk == VK_OEM_2.0 && is_ctrl_down()
}

/// 그 외 키를 WM_CHAR로 변환 시도. 실패 시 None.
fn translate_char(vk: u16, scancode: u32) -> Option<u16> {
    unsafe {
        let layout = GetKeyboardLayout(0);
        let mut state = [0u8; 256];
        if GetKeyboardState(&mut state).is_err() {
            return None;
        }
        let mut buf = [0u16; 4];
        let n = ToUnicodeEx(vk as u32, scancode, &state, &mut buf, 0, Some(layout));
        if n == 1 { Some(buf[0]) } else { None }
    }
}

unsafe fn try_fetch_more_append(list: HWND) -> bool {
    if EXHAUSTED.with(|e| e.get()) {
        return false;
    }
    let added = FETCH_MORE.with(|f| f.borrow_mut().as_mut().map(|cb| cb()).unwrap_or_default());
    if added.is_empty() {
        EXHAUSTED.with(|e| e.set(true));
        return false;
    }
    for s in &added {
        unsafe {
            push_listbox_item(list, s);
        }
    }
    true
}

unsafe fn run_pump(popup: HWND, forward_target: HWND) {
    POPUP_DONE.with(|d| d.set(false));
    unsafe {
        let list = HWND(LIST_HWND.with(|l| l.get()) as *mut c_void);
        let mut msg = MSG::default();
        while GetMessageW(&mut msg, None, 0, 0).as_bool() {
            // 우리 popup/listbox 외 다른 윈도우(특히 PostMessage로 forward한
            // HWP 윈도우)로 가는 키 메시지는 가로채지 않고 dispatch만 한다.
            // 이러지 않으면 forward한 키가 우리 pump에 다시 잡혀 무한 루프.
            let is_our_key =
                msg.message == WM_KEYDOWN && (msg.hwnd.0 == popup.0 || msg.hwnd.0 == list.0);
            if is_our_key {
                let vk = msg.wParam.0 as u16;
                let vkenum = VIRTUAL_KEY(vk);

                // cycle: ↑ / ↓ / Ctrl+/
                if vkenum == VK_UP || vkenum == VK_DOWN || is_ctrl_slash(&msg) {
                    let cnt = list_count(list);
                    if cnt > 0 {
                        let cur = list_cursel(list).max(0);
                        let next = if vkenum == VK_UP {
                            if cur > 0 { cur - 1 } else { cnt - 1 }
                        } else {
                            // ↓ 또는 Ctrl+/
                            if cur + 1 < cnt {
                                cur + 1
                            } else if try_fetch_more_append(list) {
                                cur + 1
                            } else {
                                0
                            }
                        };
                        list_setsel(list, next);
                        REPLACE_CB.with(|r| {
                            if let Some(cb) = r.borrow_mut().as_mut() {
                                cb(next as usize);
                            }
                        });
                    }
                    continue;
                }

                // Enter: 확정 + swallow
                if vkenum == VK_RETURN {
                    OUTCOME.with(|o| o.set(Outcome::Committed));
                    EXPLICIT_CLOSE.with(|e| e.set(true));
                    let _ = DestroyWindow(popup);
                    continue;
                }

                // Right / Space: 확정 + forward
                if vkenum == VK_RIGHT || vkenum == VK_SPACE {
                    OUTCOME.with(|o| o.set(Outcome::Committed));
                    EXPLICIT_CLOSE.with(|e| e.set(true));
                    let _ = DestroyWindow(popup);
                    if !forward_target.0.is_null() {
                        let _ =
                            PostMessageW(Some(forward_target), WM_KEYDOWN, msg.wParam, msg.lParam);
                        if vkenum == VK_SPACE {
                            let _ = PostMessageW(
                                Some(forward_target),
                                WM_CHAR,
                                WPARAM(b' ' as usize),
                                msg.lParam,
                            );
                        }
                        let _ =
                            PostMessageW(Some(forward_target), WM_KEYUP, msg.wParam, msg.lParam);
                    }
                    continue;
                }

                // ESC: 취소 + swallow
                if vkenum == VK_ESCAPE {
                    OUTCOME.with(|o| o.set(Outcome::Cancelled));
                    EXPLICIT_CLOSE.with(|e| e.set(true));
                    let _ = DestroyWindow(popup);
                    continue;
                }

                // 그 외 모든 키: 취소 + forward
                OUTCOME.with(|o| o.set(Outcome::Cancelled));
                EXPLICIT_CLOSE.with(|e| e.set(true));
                let _ = DestroyWindow(popup);
                if !forward_target.0.is_null() {
                    let _ = PostMessageW(Some(forward_target), WM_KEYDOWN, msg.wParam, msg.lParam);
                    let scancode = ((msg.lParam.0 >> 16) & 0xFF) as u32;
                    if let Some(ch) = translate_char(vk, scancode) {
                        let _ = PostMessageW(
                            Some(forward_target),
                            WM_CHAR,
                            WPARAM(ch as usize),
                            msg.lParam,
                        );
                    }
                    let _ = PostMessageW(Some(forward_target), WM_KEYUP, msg.wParam, msg.lParam);
                }
                continue;
            }
            DispatchMessageW(&msg);
            // popup이 destroy되면 펌프 종료. 큐에 남은 메시지(forward한 키 등)는
            // HWP main pump가 dispatch한다.
            if POPUP_DONE.with(|d| d.get()) {
                break;
            }
        }
    }
}
