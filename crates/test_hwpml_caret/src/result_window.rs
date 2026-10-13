//! 결과를 보여 주는 창. 읽기 전용 편집 컨트롤이라 글을 선택해서 복사할 수 있다.
use windows::Win32::Foundation::{HWND, LPARAM, LRESULT, WPARAM};
use windows::Win32::Graphics::Gdi::{DEFAULT_GUI_FONT, GetStockObject};
use windows::Win32::System::LibraryLoader::GetModuleHandleW;
use windows::Win32::UI::WindowsAndMessaging::{
    CW_USEDEFAULT, CreateWindowExW, DefWindowProcW, ES_AUTOVSCROLL, ES_MULTILINE, ES_READONLY,
    GW_CHILD, GetClientRect, GetWindow, MoveWindow, RegisterClassW, SendMessageW, WINDOW_EX_STYLE,
    WINDOW_STYLE, WM_SETFONT, WM_SIZE, WNDCLASSW, WS_CHILD, WS_OVERLAPPEDWINDOW, WS_VISIBLE,
    WS_VSCROLL,
};
use windows::core::{PCWSTR, w};

const CLASS_NAME: PCWSTR = w!("TestHwpmlCaretResult");

fn to_wide(s: &str) -> Vec<u16> {
    s.encode_utf16().chain(std::iter::once(0)).collect()
}

/// `text`를 새 창에 보여 준다. 창은 한글의 메시지 루프에서 돌고, 닫을 때까지 남는다.
pub fn show(title: &str, text: &str) {
    let title = to_wide(title);
    let text = to_wide(&text.replace('\n', "\r\n"));
    unsafe {
        let Ok(module) = GetModuleHandleW(None) else {
            return;
        };
        let instance = module.into();
        // 두 번째 호출부터는 이미 등록돼 있어 실패하지만 그대로 쓰면 된다.
        RegisterClassW(&WNDCLASSW {
            lpfnWndProc: Some(wnd_proc),
            hInstance: instance,
            lpszClassName: CLASS_NAME,
            ..Default::default()
        });
        let Ok(frame) = CreateWindowExW(
            WINDOW_EX_STYLE(0),
            CLASS_NAME,
            PCWSTR(title.as_ptr()),
            WS_OVERLAPPEDWINDOW | WS_VISIBLE,
            CW_USEDEFAULT,
            CW_USEDEFAULT,
            900,
            600,
            None,
            None,
            Some(instance),
            None,
        ) else {
            return;
        };
        let style = WS_CHILD
            | WS_VISIBLE
            | WS_VSCROLL
            | WINDOW_STYLE((ES_MULTILINE | ES_READONLY | ES_AUTOVSCROLL) as u32);
        let Ok(edit) = CreateWindowExW(
            WINDOW_EX_STYLE(0),
            w!("EDIT"),
            PCWSTR(text.as_ptr()),
            style,
            0,
            0,
            0,
            0,
            Some(frame),
            None,
            Some(instance),
            None,
        ) else {
            return;
        };
        let font = GetStockObject(DEFAULT_GUI_FONT);
        SendMessageW(edit, WM_SETFONT, Some(WPARAM(font.0 as usize)), Some(LPARAM(1)));
        fit_child(frame);
    }
}

/// 편집 컨트롤을 창의 클라이언트 영역 전체로 맞춘다.
unsafe fn fit_child(frame: HWND) {
    unsafe {
        let mut rect = Default::default();
        if let Ok(edit) = GetWindow(frame, GW_CHILD)
            && GetClientRect(frame, &mut rect).is_ok()
        {
            let _ = MoveWindow(edit, 0, 0, rect.right, rect.bottom, true);
        }
    }
}

unsafe extern "system" fn wnd_proc(hwnd: HWND, msg: u32, wparam: WPARAM, lparam: LPARAM) -> LRESULT {
    unsafe {
        if msg == WM_SIZE {
            fit_child(hwnd);
        }
        DefWindowProcW(hwnd, msg, wparam, lparam)
    }
}
