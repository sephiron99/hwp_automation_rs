//! 가로챈 키 입력을 다른 윈도우(HWP 메인)로 전달하는 헬퍼.
//!
//! `PostMessageW`로 임의 메시지를 보내는 것은 winsafe에서도 unsafe로
//! 분류되는 작업이라, `#![forbid(unsafe_code)]`인 플러그인 크레이트 대신
//! FFI 계층인 이 크레이트에 safe API로 감싸서 둔다.

use std::ffi::c_void;

use windows::Win32::Foundation::{HWND, LPARAM, WPARAM};
use windows::Win32::UI::WindowsAndMessaging::{PostMessageW, WM_CHAR, WM_KEYDOWN, WM_KEYUP};

/// 키 메시지 시퀀스를 `target` 윈도우로 post한다: KEYDOWN → [CHAR] → KEYUP.
///
/// - `target` — 대상 HWND 값(`HWND::ptr() as usize` 등). 0이면 no-op.
/// - `key_wparam`/`key_lparam` — 가로챈 WM_KEYDOWN의 wParam(가상 키)/lParam.
///   KEYUP에도 같은 lParam을 재사용한다(기존 popup forward 동작과 동일).
/// - `char_wparam` — 해당 키의 WM_CHAR wParam(문자 코드). 있으면 KEYDOWN과
///   KEYUP 사이에 post.
pub fn forward_key(target: usize, key_wparam: usize, key_lparam: isize, char_wparam: Option<usize>) {
    if target == 0 {
        return;
    }
    let hwnd = HWND(target as *mut c_void);
    let wp = WPARAM(key_wparam);
    let lp = LPARAM(key_lparam);
    // SAFETY: PostMessageW는 메시지를 큐에 넣기만 하며, 키 메시지의
    // wParam/lParam은 수신 측에서 값으로만 해석된다(포인터 아님).
    unsafe {
        let _ = PostMessageW(Some(hwnd), WM_KEYDOWN, wp, lp);
        if let Some(c) = char_wparam {
            let _ = PostMessageW(Some(hwnd), WM_CHAR, WPARAM(c), lp);
        }
        let _ = PostMessageW(Some(hwnd), WM_KEYUP, wp, lp);
    }
}
