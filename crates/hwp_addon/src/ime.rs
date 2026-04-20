//! IME(입력기) 조합 문자열 제어 유틸리티.
//!
//! 단축키로 애드온 액션이 실행될 때 한글 입력기가 조합(composition) 상태라면
//! 후속 액션이 조합 중인 문자열을 넘지 못해 커서 조작이 꼬이는 문제가 있다
//! (툴바 클릭 시에는 focus-out 부수효과로 IME가 자동 commit 되어 문제가
//! 발생하지 않음).
//!
//! # 진단 이력
//! - HWP 2022/2024의 `HwpMainEditWnd`는 `ImmGetContext`에 NULL을 반환.
//! - 부모 창 클래스가 `HwndWrapper[Hwp.exe;;<guid>]` — WPF `HwndSource` 시그니처.
//! - WPF 측 IMC는 `comp_bytes=0, open=false`로 IMM32 레벨에 조합이 안 보임.
//! - 결론: HWP는 TSF(Text Services Framework)로 IME를 처리하며 IMM32는 무력.
//!
//! # 시도 순서
//! [`commit_composition`]은 효과가 큰 순서대로 여러 방법을 시도한다:
//! 1. **SetFocus 해킹** — 부모 HwndWrapper로 잠시 포커스를 옮겼다가 복귀.
//!    툴바 클릭과 동일한 부수효과(WM_KILLFOCUS → IME commit)를 기대.
//! 2. **TSF `TerminateComposition`** — `ITfThreadMgr::GetFocus()`로 현재 포커스
//!    `ITfDocumentMgr`를 얻고, top `ITfContext`에서 `ITfContextOwnerCompositionServices::TerminateComposition(NULL)`.
//! 3. **IMM32 fallback + 진단 로그** — 여러 HWND 후보에서 IMC를 찾아 시도.
//!    실제로는 효과가 없지만 환경 변화 시 감지용으로 남김.

use crate::debug::log;
use windows::Win32::Foundation::{HWND, LPARAM};
use windows::Win32::System::Com::{CLSCTX_INPROC_SERVER, CoCreateInstance};
use windows::Win32::UI::Input::Ime::{
    CPS_COMPLETE, GCS_COMPSTR, ImmGetCompositionStringW, ImmGetContext, ImmGetDefaultIMEWnd,
    ImmGetOpenStatus, ImmNotifyIME, ImmReleaseContext, NI_COMPOSITIONSTR,
};
use windows::Win32::UI::Input::KeyboardAndMouse::{GetFocus, SetFocus};
use windows::Win32::UI::TextServices::{
    CLSID_TF_ThreadMgr, ITfContext, ITfContextOwnerCompositionServices, ITfDocumentMgr,
    ITfThreadMgr,
};
use windows::Win32::UI::WindowsAndMessaging::{
    EnumChildWindows, GA_PARENT, GA_ROOT, GA_ROOTOWNER, GetAncestor, GetClassNameW,
};
use windows::core::{BOOL, Interface};

// ============================================================================
// Public API
// ============================================================================

/// 현재 스레드의 IME 조합을 확정(또는 취소)시켜 문서에 반영되도록 한다.
///
/// HWP가 TSF를 쓰므로 TSF 경로가 핵심이지만, 실패 시 부수적으로 SetFocus와
/// IMM32를 시도한다. 어느 한 경로가 효과를 내면 성공.
///
/// 이 함수는 절대 실패로 호출측을 중단시키지 않는다. 모든 단계에서 오류는
/// 로그로만 남기고 조용히 다음 단계로 진행한다.
pub fn commit_composition() {
    unsafe {
        log("ime", "commit_composition: 시작");

        // 1. TSF 경로 — 가장 정통. 단계별 로그는 함수 내부에서 직접 남김.
        terminate_tsf_composition();

        // 2. SetFocus 해킹 — 툴바 클릭과 동일한 focus-out 부수효과
        try_setfocus_bounce();

        // 3. IMM32 경로 — 진단용 (실제로는 효과 없음이 확인됨)
        try_imm32_all_candidates();

        log("ime", "commit_composition: 종료");
    }
}

// ============================================================================
// TSF 경로
// ============================================================================

/// TSF `ITfThreadMgr`를 통해 현재 포커스 문서의 모든 composition을 종료한다.
///
/// 단계별로 세분화된 에러 메시지를 로그에 남겨, 실패가 어느 단계에서 났는지
/// (CoCreateInstance / GetFocus / GetTop / QI / TerminateComposition) 즉시
/// 구분할 수 있게 한다. `windows::core::Error`가 S_OK(0x00000000)와 NULL
/// out-pointer를 "success with null"로 래핑해 에러로 보고하는 케이스가
/// 있어(예: GetFocus — TSF가 아직 활성화되지 않은 스레드), 단계별 로깅이
/// 진단에 필수적이다.
unsafe fn terminate_tsf_composition() {
    unsafe {
        let thread_mgr: ITfThreadMgr =
            match CoCreateInstance(&CLSID_TF_ThreadMgr, None, CLSCTX_INPROC_SERVER) {
                Ok(tm) => tm,
                Err(e) => {
                    log("ime", &format!("TSF: CoCreateInstance 실패: {e}"));
                    return;
                }
            };

        let doc_mgr: ITfDocumentMgr = match thread_mgr.GetFocus() {
            Ok(dm) => dm,
            Err(e) => {
                // S_OK + NULL도 여기로 들어온다: TSF가 이 스레드에 아직
                // 초기화되지 않았거나 focus document가 없는 상태.
                log("ime", &format!("TSF: GetFocus 실패: {e}"));
                return;
            }
        };

        let ctx: ITfContext = match doc_mgr.GetTop() {
            Ok(c) => c,
            Err(e) => {
                log("ime", &format!("TSF: GetTop 실패: {e}"));
                return;
            }
        };

        let cs: ITfContextOwnerCompositionServices = match ctx.cast() {
            Ok(s) => s,
            Err(e) => {
                log("ime", &format!("TSF: QI ContextOwnerCompositionServices 실패: {e}"));
                return;
            }
        };

        match cs.TerminateComposition(None) {
            Ok(()) => log("ime", "TSF: TerminateComposition 성공"),
            Err(e) => log("ime", &format!("TSF: TerminateComposition 실패: {e}")),
        }
    }
}

// ============================================================================
// SetFocus 해킹
// ============================================================================

/// 포커스를 부모 창으로 잠깐 옮겼다가 다시 원래 창으로 복귀시킨다.
///
/// 툴바 클릭이 문제를 일으키지 않는 이유는 focus-out 시점에 HWP 편집창이
/// 내부적으로 조합을 커밋하기 때문이다. 그 부수효과를 강제로 재현한다.
unsafe fn try_setfocus_bounce() {
    unsafe {
        let focus = GetFocus();
        if focus.is_invalid() {
            log("ime", "setfocus: GetFocus=null, 스킵");
            return;
        }
        let parent = GetAncestor(focus, GA_PARENT);
        if parent.is_invalid() {
            log("ime", "setfocus: parent=null, 스킵");
            return;
        }
        log(
            "ime",
            &format!("setfocus: {focus:?} → {parent:?} → {focus:?}"),
        );
        let _ = SetFocus(Some(parent));
        let _ = SetFocus(Some(focus));
    }
}

// ============================================================================
// IMM32 fallback (진단용)
// ============================================================================

unsafe fn class_name_of(hwnd: HWND) -> String {
    unsafe {
        let mut buf = [0u16; 128];
        let len = GetClassNameW(hwnd, &mut buf);
        if len > 0 {
            String::from_utf16_lossy(&buf[..len as usize])
        } else {
            String::from("?")
        }
    }
}

unsafe fn try_imm32_commit(hwnd: HWND, tag: &str) -> bool {
    unsafe {
        if hwnd.is_invalid() {
            return false;
        }
        let himc = ImmGetContext(hwnd);
        if himc.0.is_null() {
            return false;
        }
        let open = ImmGetOpenStatus(himc).as_bool();
        let comp_bytes = ImmGetCompositionStringW(himc, GCS_COMPSTR, None, 0);
        // 조합이 실제로 있는 IMC만 로그로 남긴다 (0인 경우는 노이즈).
        if comp_bytes > 0 {
            let class_name = class_name_of(hwnd);
            log(
                "ime",
                &format!(
                    "[{tag}] hwnd={hwnd:?} class={class_name} IMC={:?} open={open} comp_bytes={comp_bytes}",
                    himc.0
                ),
            );
            let ok = ImmNotifyIME(himc, NI_COMPOSITIONSTR, CPS_COMPLETE, 0).as_bool();
            log("ime", &format!("[{tag}] ImmNotifyIME(CPS_COMPLETE) -> {ok}"));
            let _ = ImmReleaseContext(hwnd, himc);
            return ok;
        }
        let _ = ImmReleaseContext(hwnd, himc);
        false
    }
}

extern "system" fn enum_child_proc(hwnd: HWND, _lparam: LPARAM) -> BOOL {
    unsafe {
        let _ = try_imm32_commit(hwnd, "child");
    }
    BOOL(1)
}

unsafe fn try_imm32_all_candidates() {
    unsafe {
        let focus = GetFocus();
        if focus.is_invalid() {
            return;
        }
        try_imm32_commit(focus, "focus");
        try_imm32_commit(GetAncestor(focus, GA_PARENT), "parent");
        let root = GetAncestor(focus, GA_ROOT);
        try_imm32_commit(root, "root");
        try_imm32_commit(GetAncestor(focus, GA_ROOTOWNER), "rootOwner");
        try_imm32_commit(ImmGetDefaultIMEWnd(focus), "defaultIME");
        if !root.is_invalid() {
            let _ = EnumChildWindows(Some(root), Some(enum_child_proc), LPARAM(0));
        }
    }
}
