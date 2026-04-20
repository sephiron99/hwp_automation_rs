//! HWP 창 선택 UI (Win32 전용).
//!
//! 현재 실행 중인 HWP 창을 나열하고, 사용자가 기존 창 중 하나를 고르거나
//! 새 인스턴스를 생성하도록 한다.
//!
//! 기존 창 선택 시 `AccessibleObjectFromWindow(OBJID_NATIVEOM)` (MSAA)로
//! `IDispatch`를 가져오려 시도한다. Office 계열이 쓰는 표준 경로이며, HWP가
//! 지원하면 동작한다. 지원하지 않는 경우 호출자 쪽에서 에러를 처리하고
//! 다시 UI를 띄우거나 새 인스턴스 경로로 fallback 하면 된다.

use std::cell::{Cell, RefCell};
use std::ffi::c_void;

use hwp_core::error::{HwpError, Result};
use hwp_core::hwp_obj::HwpObject;
use windows::core::{w, Interface};
use windows::Win32::Foundation::{BOOL, HINSTANCE, HWND, LPARAM, LRESULT, TRUE, WPARAM};
use windows::Win32::Graphics::Gdi::HBRUSH;
use windows::Win32::System::Com::IDispatch;
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::Accessibility::AccessibleObjectFromWindow;
use windows::Win32::UI::WindowsAndMessaging::{
    CreateWindowExW, DefWindowProcW, DestroyWindow, DispatchMessageW, EnumWindows,
    GetClassNameW, GetMessageW, GetWindowTextW, IsWindowVisible, LoadCursorW, PostQuitMessage,
    RegisterClassW, SendMessageW, ShowWindow, TranslateMessage, UpdateWindow, BS_DEFPUSHBUTTON,
    COLOR_WINDOW, CW_USEDEFAULT, HMENU, IDC_ARROW, LB_ADDSTRING, LB_GETCURSEL, LB_SETCURSEL,
    LBS_NOTIFY, MSG, SW_SHOW, WINDOW_EX_STYLE, WINDOW_STYLE, WM_CLOSE, WM_COMMAND, WM_CREATE,
    WM_DESTROY, WNDCLASSW, WS_BORDER, WS_CHILD, WS_OVERLAPPEDWINDOW, WS_VISIBLE, WS_VSCROLL,
};

/// 사용자가 선택한 결과.
#[derive(Clone, Debug)]
pub enum Choice {
    /// 기존 HWP 창 선택.
    Existing(HWND),
    /// 새 HWP 인스턴스 생성.
    New,
    /// 취소.
    Cancelled,
}

// wndproc은 캡처를 못 하므로 단일 스레드 내에서 thread_local로 상태를 공유한다.
thread_local! {
    static ENTRIES: RefCell<Vec<(HWND, String)>> = const { RefCell::new(Vec::new()) };
    static CHOICE: RefCell<Choice> = const { RefCell::new(Choice::Cancelled) };
    static LIST_HWND: Cell<isize> = const { Cell::new(0) };
}

const ID_LIST: isize = 101;
const ID_OK: isize = 102;
const ID_CANCEL: isize = 103;

unsafe extern "system" fn enum_hwp(hwnd: HWND, _: LPARAM) -> BOOL {
    unsafe {
        if !IsWindowVisible(hwnd).as_bool() {
            return TRUE;
        }
        let mut class_buf = [0u16; 128];
        let n = GetClassNameW(hwnd, &mut class_buf);
        if n == 0 {
            return TRUE;
        }
        let cls = String::from_utf16_lossy(&class_buf[..n as usize]);
        // HWP 창 클래스명은 일반적으로 "Hwp"로 시작한다 (예: HwpFrameWindow).
        if !cls.to_lowercase().starts_with("hwp") {
            return TRUE;
        }
        let mut title_buf = [0u16; 512];
        let tn = GetWindowTextW(hwnd, &mut title_buf);
        let title = if tn > 0 {
            String::from_utf16_lossy(&title_buf[..tn as usize])
        } else {
            String::from("(제목 없음)")
        };
        let display = format!("{title}  [{cls}]");
        ENTRIES.with(|e| e.borrow_mut().push((hwnd, display)));
        TRUE
    }
}

unsafe extern "system" fn wnd_proc(hwnd: HWND, msg: u32, wp: WPARAM, lp: LPARAM) -> LRESULT {
    unsafe {
        match msg {
            WM_CREATE => {
                on_create(hwnd);
                LRESULT(0)
            }
            WM_COMMAND => {
                let id = (wp.0 & 0xFFFF) as isize;
                match id {
                    ID_OK => handle_ok(hwnd),
                    ID_CANCEL => {
                        CHOICE.with(|c| *c.borrow_mut() = Choice::Cancelled);
                        let _ = DestroyWindow(hwnd);
                    }
                    _ => {}
                }
                LRESULT(0)
            }
            WM_CLOSE => {
                CHOICE.with(|c| *c.borrow_mut() = Choice::Cancelled);
                let _ = DestroyWindow(hwnd);
                LRESULT(0)
            }
            WM_DESTROY => {
                PostQuitMessage(0);
                LRESULT(0)
            }
            _ => DefWindowProcW(hwnd, msg, wp, lp),
        }
    }
}

unsafe fn handle_ok(hwnd: HWND) {
    unsafe {
        let list = HWND(LIST_HWND.with(|l| l.get()) as *mut c_void);
        let sel = SendMessageW(list, LB_GETCURSEL, Some(WPARAM(0)), Some(LPARAM(0))).0;
        let choice = if sel < 0 {
            Choice::Cancelled
        } else if sel == 0 {
            Choice::New
        } else {
            ENTRIES.with(|e| {
                let entries = e.borrow();
                let idx = (sel - 1) as usize;
                entries
                    .get(idx)
                    .map(|(h, _)| Choice::Existing(*h))
                    .unwrap_or(Choice::Cancelled)
            })
        };
        CHOICE.with(|c| *c.borrow_mut() = choice);
        let _ = DestroyWindow(hwnd);
    }
}

unsafe fn on_create(hwnd: HWND) {
    unsafe {
        let hmodule = GetModuleHandleW(None).unwrap();
        let hinstance = HINSTANCE(hmodule.0);

        // Listbox
        let list = CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            w!("LISTBOX"),
            w!(""),
            WS_CHILD | WS_VISIBLE | WS_BORDER | WS_VSCROLL | WINDOW_STYLE(LBS_NOTIFY as u32),
            10,
            10,
            460,
            280,
            Some(hwnd),
            Some(HMENU(ID_LIST as *mut c_void)),
            Some(hinstance),
            None,
        )
        .unwrap();
        LIST_HWND.with(|l| l.set(list.0 as isize));

        let push = |text: &str| {
            let wide: Vec<u16> = text.encode_utf16().chain(std::iter::once(0)).collect();
            SendMessageW(
                list,
                LB_ADDSTRING,
                Some(WPARAM(0)),
                Some(LPARAM(wide.as_ptr() as isize)),
            );
        };
        push("[+] 새 HWP 창 만들기");
        ENTRIES.with(|e| {
            for (_, display) in e.borrow().iter() {
                push(display);
            }
        });
        SendMessageW(list, LB_SETCURSEL, Some(WPARAM(0)), Some(LPARAM(0)));

        // OK 버튼
        CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            w!("BUTTON"),
            w!("확인"),
            WS_CHILD | WS_VISIBLE | WINDOW_STYLE(BS_DEFPUSHBUTTON as u32),
            300,
            305,
            80,
            28,
            Some(hwnd),
            Some(HMENU(ID_OK as *mut c_void)),
            Some(hinstance),
            None,
        )
        .unwrap();

        // 취소 버튼
        CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            w!("BUTTON"),
            w!("취소"),
            WS_CHILD | WS_VISIBLE,
            390,
            305,
            80,
            28,
            Some(hwnd),
            Some(HMENU(ID_CANCEL as *mut c_void)),
            Some(hinstance),
            None,
        )
        .unwrap();
    }
}

/// 대화상자를 띄우고 사용자의 선택을 반환한다.
///
/// 메시지 루프는 창이 닫힐 때까지 블록한다.
pub fn show_dialog() -> Choice {
    // 열린 HWP 창 수집
    ENTRIES.with(|e| e.borrow_mut().clear());
    unsafe {
        let _ = EnumWindows(Some(enum_hwp), LPARAM(0));
    }
    CHOICE.with(|c| *c.borrow_mut() = Choice::Cancelled);

    unsafe {
        let hmodule = GetModuleHandleW(None).unwrap();
        let hinstance = HINSTANCE(hmodule.0);
        let class_name = w!("HwpDabbrevDialog");

        let wc = WNDCLASSW {
            hCursor: LoadCursorW(None, IDC_ARROW).unwrap_or_default(),
            hInstance: hinstance,
            lpszClassName: class_name,
            lpfnWndProc: Some(wnd_proc),
            // "시스템 창 색상 + 1" 관용구 — WNDCLASS의 hbrBackground 해석 규칙.
            hbrBackground: HBRUSH((COLOR_WINDOW.0 as usize + 1) as *mut c_void),
            ..Default::default()
        };
        RegisterClassW(&wc);

        let hwnd = CreateWindowExW(
            WINDOW_EX_STYLE::default(),
            class_name,
            w!("HWP 창 선택 - hwp_dabbrev"),
            WS_OVERLAPPEDWINDOW,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            500,
            380,
            None,
            None,
            Some(hinstance),
            None,
        )
        .unwrap();

        let _ = ShowWindow(hwnd, SW_SHOW);
        let _ = UpdateWindow(hwnd);

        let mut msg = MSG::default();
        while GetMessageW(&mut msg, None, 0, 0).as_bool() {
            let _ = TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    }

    CHOICE.with(|c| c.borrow().clone())
}

/// HWP 창의 HWND에서 `IDispatch`를 얻어 `HwpObject`를 생성한다.
///
/// MSAA의 `OBJID_NATIVEOM` 경로를 사용한다. HWP가 이 인터페이스를 노출하지
/// 않으면 실패한다.
pub fn hwp_from_hwnd(hwnd: HWND) -> Result<HwpObject> {
    const OBJID_NATIVEOM: u32 = 0xFFFF_FFF0;
    unsafe {
        let mut dispatch: Option<IDispatch> = None;
        AccessibleObjectFromWindow(
            hwnd,
            OBJID_NATIVEOM,
            &IDispatch::IID,
            &mut dispatch as *mut _ as *mut *mut c_void,
        )?;
        let dispatch = dispatch.ok_or(HwpError::ConnectionFailed)?;
        HwpObject::new(dispatch)
    }
}
