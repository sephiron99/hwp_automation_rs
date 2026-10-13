use windows::Win32::UI::WindowsAndMessaging::{MessageBoxW, MB_ICONERROR, MB_ICONINFORMATION, MB_OK};
use windows::core::PCWSTR;

fn to_wide(s: &str) -> Vec<u16> {
    s.encode_utf16().chain(std::iter::once(0)).collect()
}

/// 정보 메시지 박스를 표시합니다.
pub fn msgbox(title: &str, msg: &str) {
    tracing::info!(title, msg, "msgbox");
    let title = to_wide(title);
    let msg = to_wide(msg);
    unsafe {
        MessageBoxW(None, PCWSTR(msg.as_ptr()), PCWSTR(title.as_ptr()), MB_OK | MB_ICONINFORMATION);
    }
}

/// 에러 메시지 박스를 표시합니다.
pub fn msgbox_err(title: &str, msg: &str) {
    tracing::error!(title, msg, "msgbox_err");
    let title = to_wide(title);
    let msg = to_wide(msg);
    unsafe {
        MessageBoxW(None, PCWSTR(msg.as_ptr()), PCWSTR(title.as_ptr()), MB_OK | MB_ICONERROR);
    }
}
