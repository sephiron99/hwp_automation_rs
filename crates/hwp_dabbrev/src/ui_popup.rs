//! Dabbrev 자동완성 popup (Win32 전용, winsafe gui 기반).
//!
//! 자식 컨트롤 없이 후보 목록을 직접 그리는(owner-drawn) WS_POPUP 창.
//! winsafe `WindowMain::run_main()`이 동기 message pump를 돌리고, 창이
//! 파괴되면(내부 WM_NCDESTROY → PostQuitMessage) WM_QUIT를 이 중첩 펌프가
//! 소비하고 리턴하므로 HWP 메인 펌프에는 영향이 없다.
//!
//! ## 키 매핑
//! - ↑ / ↓ / Ctrl+/ — preview replace + popup 유지 (cycle)
//! - Enter — 확정, popup 닫음, 키 swallow
//! - Right / Space — 확정, popup 닫음, 키를 HWP로 forward
//! - ESC — 취소, popup 닫음, 키 swallow
//! - 그 외 모든 키 — 취소, popup 닫음, 키를 HWP로 forward
//!
//! 키→문자 변환은 ToUnicodeEx 대신, 펌프의 TranslateMessage가 이미 큐에
//! 넣어 둔 WM_CHAR을 keydown 처리 중 PeekMessage로 회수하는 방식을 쓴다.
//! forward 자체는 펌프 종료 후 `hwp_addon::keyfwd`로 수행하므로, forward한
//! 키를 우리 펌프가 다시 가로채는 문제가 원천적으로 없다.

use std::cell::{Cell, RefCell};
use std::panic::{AssertUnwindSafe, catch_unwind};
use std::rc::Rc;

use hwp_addon::debug::log;
use winsafe::{self as w, co, gui, msg::WndMsg, prelude::*};

const POPUP_WIDTH: i32 = 200;
const POPUP_HEIGHT: i32 = 120;

/// popup 종료 원인.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Outcome {
    /// Enter / Right / Space — 명시적 확정.
    Committed,
    /// ESC / 그 외 키 / 닫힘 — 취소.
    Cancelled,
}

/// 종료 후 HWP로 forward할 키 — 가로챈 WM_KEYDOWN의 raw wParam/lParam과,
/// TranslateMessage가 생성해 둔 WM_CHAR의 wParam(있을 때만).
struct PendingKey {
    wparam: usize,
    lparam: isize,
    char_wparam: Option<usize>,
}

/// popup 한 세션의 상태. 이벤트 클로저들이 `Rc`로 공유하며, 가변성은 전부
/// Cell/RefCell(winsafe 이벤트 클로저는 `Fn`이라 내부 가변성 필수).
struct Session {
    items: RefCell<Vec<String>>,
    sel: Cell<usize>,
    /// 스크롤 오프셋 — 첫 표시 행의 item index.
    top: Cell<usize>,
    /// wm_paint에서 실측한 행 높이(px). 첫 paint 전까지 0.
    row_h: Cell<i32>,
    outcome: Cell<Outcome>,
    /// 처음 WM_ACTIVATE(WA_ACTIVE)를 받은 후에만 deactivate를 close 트리거로
    /// 본다. 초기 생성 시점의 transient deactivate에 즉시 닫히지 않도록.
    got_active: Cell<bool>,
    /// 키 분기에서 DestroyWindow 직전에 set. DestroyWindow가 동기 발생시키는
    /// WM_ACTIVATE(WA_INACTIVE)가 outcome을 덮어쓰지 못하도록 가드.
    explicit_close: Cell<bool>,
    exhausted: Cell<bool>,
    fetch_more: RefCell<Box<dyn FnMut() -> Vec<String>>>,
    replace: RefCell<Box<dyn FnMut(usize)>>,
    pending: RefCell<Option<PendingKey>>,
}

/// popup을 띄우고 사용자 조작이 끝날 때까지 블록한다.
///
/// - `initial` — 첫 batch candidates. `start_index`는 첫 sel 위치 (preview
///   replace는 이미 호출자가 수행했음).
/// - `fetch_more` — 목록 끝에서 사용자가 더 내려갈 때 호출. 빈 Vec 반환 시
///   더 호출하지 않음.
/// - `replace` — sel이 바뀔 때마다 호출. 호출자가 문서에 preview replace 수행.
/// - `forward_target` — 확정/취소 후 키를 forward할 윈도우 (HWP 메인).
pub fn show(
    initial: &[String],
    start_index: usize,
    fetch_more: impl FnMut() -> Vec<String> + 'static,
    replace: impl FnMut(usize) + 'static,
    forward_target: Option<w::HWND>,
) -> Outcome {
    // hook이 단축키를 가로채지 않도록 양보 — popup이 직접 Ctrl+/ 처리.
    hwp_addon::shortcut::set_popup_active(true);

    // run_main은 창 생성 실패 시 panic한다. FFI 경계(HWP의 DoAction 호출)를
    // unwind가 넘지 않도록 여기서 잡아 Cancelled로 처리.
    let (outcome, pending) =
        catch_unwind(AssertUnwindSafe(|| run_popup(initial, start_index, fetch_more, replace)))
            .unwrap_or_else(|_| {
                log("ui_popup", "popup panic — cancelled 처리");
                (Outcome::Cancelled, None)
            });

    if let (Some(target), Some(k)) = (&forward_target, pending) {
        hwp_addon::keyfwd::forward_key(target.ptr() as usize, k.wparam, k.lparam, k.char_wparam);
    }

    // popup destroy 후 HWP를 다시 활성화. ShowWindow 계열은 절대 호출하지
    // 않는다 — SW_RESTORE는 maximize 상태를 normal로 되돌리면서 위치를
    // 옮기고, SW_SHOWNORMAL도 동일한 부작용이 있다. BringWindowToTop과
    // SetForegroundWindow는 z-order/activation만 바꾸므로 안전.
    if let Some(target) = &forward_target {
        let _ = target.BringWindowToTop();
        let _ = target.SetForegroundWindow();
    }

    hwp_addon::shortcut::set_popup_active(false);
    outcome
}

/// 창 생성 + 이벤트 등록 + 동기 펌프. 종료 시 (outcome, forward할 키) 반환.
fn run_popup(
    initial: &[String],
    start_index: usize,
    fetch_more: impl FnMut() -> Vec<String> + 'static,
    replace: impl FnMut(usize) + 'static,
) -> (Outcome, Option<PendingKey>) {
    let anchor = try_anchor();
    let se = Rc::new(Session {
        items: RefCell::new(initial.to_vec()),
        sel: Cell::new(start_index.min(initial.len().saturating_sub(1))),
        top: Cell::new(0),
        row_h: Cell::new(0),
        outcome: Cell::new(Outcome::Cancelled),
        got_active: Cell::new(false),
        explicit_close: Cell::new(false),
        exhausted: Cell::new(false),
        fetch_more: RefCell::new(Box::new(fetch_more)),
        replace: RefCell::new(Box::new(replace)),
        pending: RefCell::new(None),
    });

    let wnd = gui::WindowMain::new(gui::WindowMainOpts {
        class_name: "HwpDabbrevPopup",
        title: "dabbrev",
        size: (POPUP_WIDTH, POPUP_HEIGHT),
        style: co::WS::POPUP | co::WS::BORDER,
        ex_style: co::WS_EX::TOPMOST | co::WS_EX::TOOLWINDOW,
        class_bg_brush: gui::Brush::Color(co::COLOR::WINDOW),
        // IsDialogMessage가 화살표/Enter/ESC를 가로채지 않고 wndproc까지
        // 그대로 오도록 한다.
        process_dlg_msgs: false,
        ..Default::default()
    });

    // WindowMainOpts에는 위치 지정이 없어(화면 중앙 고정) 표시 전에 옮긴다.
    let wnd2 = wnd.clone();
    wnd.on().wm_create(move |_| {
        let _ = wnd2.hwnd().SetWindowPos(
            w::HwndPlace::None,
            anchor,
            w::SIZE::default(),
            co::SWP::NOSIZE | co::SWP::NOZORDER | co::SWP::NOACTIVATE,
        );
        Ok(0)
    });

    let wnd2 = wnd.clone();
    let se2 = se.clone();
    wnd.on().wm(co::WM::KEYDOWN, move |p: WndMsg| {
        on_key_down(&se2, wnd2.hwnd(), &p);
        Ok(0) // handled — DefWindowProc로 가지 않음 (swallow)
    });

    let wnd2 = wnd.clone();
    let se2 = se.clone();
    wnd.on().wm_activate(move |p| {
        if p.event == co::WA::INACTIVE {
            // EXPLICIT_CLOSE: 키 분기에서 우리가 직접 DestroyWindow한 경우
            // — 이미 outcome이 정해졌으므로 건드리지 않는다.
            if se2.got_active.get() && !se2.explicit_close.get() {
                // 외부 요인(Alt+Tab, 외부 클릭 등)으로 활성 잃음 → cancel.
                se2.outcome.set(Outcome::Cancelled);
                let _ = wnd2.hwnd().DestroyWindow();
            }
        } else {
            se2.got_active.set(true);
        }
        Ok(())
    });

    let wnd2 = wnd.clone();
    let se2 = se.clone();
    wnd.on().wm_paint(move || {
        paint(&se2, wnd2.hwnd())?;
        Ok(())
    });

    let wnd2 = wnd.clone();
    let se2 = se.clone();
    wnd.on().wm_l_button_down(move |p| {
        let row_h = se2.row_h.get().max(1);
        let idx = se2.top.get() + (p.coords.y.max(0) / row_h) as usize;
        if idx < se2.items.borrow().len() && idx != se2.sel.get() {
            se2.sel.set(idx);
            (se2.replace.borrow_mut())(idx);
            let _ = wnd2.hwnd().InvalidateRect(None, true);
        }
        Ok(())
    });

    if let Err(e) = wnd.run_main(None) {
        log("ui_popup", &format!("run_main 오류: {e}"));
    }

    (se.outcome.get(), se.pending.borrow_mut().take())
}

/// WM_KEYDOWN 분기. 모든 키는 swallow되며, forward가 필요한 키는 pending에
/// 기록해 펌프 종료 후 보낸다.
fn on_key_down(se: &Session, hwnd: &w::HWND, p: &WndMsg) {
    let vk = p.wparam as u16;
    let ctrl_slash = vk == co::VK::OEM_2.raw() && w::GetAsyncKeyState(co::VK::CONTROL);

    // cycle: ↑ / ↓ / Ctrl+/
    if vk == co::VK::UP.raw() || vk == co::VK::DOWN.raw() || ctrl_slash {
        let _ = drain_char(hwnd);
        cycle(se, hwnd, vk == co::VK::UP.raw());
        return;
    }

    // Enter: 확정 + swallow
    if vk == co::VK::RETURN.raw() {
        let _ = drain_char(hwnd); // '\r' 버림
        close(se, hwnd, Outcome::Committed, None);
        return;
    }

    // Right / Space: 확정 + forward (Space의 ' '는 WM_CHAR로 자연 회수됨)
    if vk == co::VK::RIGHT.raw() || vk == co::VK::SPACE.raw() {
        let ch = drain_char(hwnd);
        close(se, hwnd, Outcome::Committed, Some(PendingKey {
            wparam: p.wparam,
            lparam: p.lparam,
            char_wparam: ch,
        }));
        return;
    }

    // ESC: 취소 + swallow
    if vk == co::VK::ESCAPE.raw() {
        let _ = drain_char(hwnd);
        close(se, hwnd, Outcome::Cancelled, None);
        return;
    }

    // 그 외 모든 키: 취소 + forward
    let ch = drain_char(hwnd);
    close(se, hwnd, Outcome::Cancelled, Some(PendingKey {
        wparam: p.wparam,
        lparam: p.lparam,
        char_wparam: ch,
    }));
}

fn close(se: &Session, hwnd: &w::HWND, outcome: Outcome, pending: Option<PendingKey>) {
    se.outcome.set(outcome);
    se.explicit_close.set(true);
    *se.pending.borrow_mut() = pending;
    let _ = hwnd.DestroyWindow();
}

/// 직전 WM_KEYDOWN을 TranslateMessage가 변환해 큐에 넣은 WM_CHAR이 있으면
/// 꺼내서 wParam(문자 코드)을 반환한다. 없으면(화살표 등 비문자 키) None.
fn drain_char(hwnd: &w::HWND) -> Option<usize> {
    let mut msg = w::MSG::default();
    let has = w::PeekMessage(
        &mut msg,
        Some(hwnd),
        co::WM::CHAR.raw(),
        co::WM::CHAR.raw(),
        co::PM::REMOVE,
    );
    if has { Some(msg.wParam) } else { None }
}

fn cycle(se: &Session, hwnd: &w::HWND, up: bool) {
    let count = se.items.borrow().len();
    if count == 0 {
        return;
    }
    let cur = se.sel.get();
    let next = if up {
        if cur > 0 { cur - 1 } else { count - 1 }
    } else {
        // ↓ 또는 Ctrl+/
        if cur + 1 < count {
            cur + 1
        } else if try_fetch_more_append(se) {
            cur + 1
        } else {
            0
        }
    };
    se.sel.set(next);
    (se.replace.borrow_mut())(next);
    let _ = hwnd.InvalidateRect(None, true);
}

fn try_fetch_more_append(se: &Session) -> bool {
    if se.exhausted.get() {
        return false;
    }
    let added = (se.fetch_more.borrow_mut())();
    if added.is_empty() {
        se.exhausted.set(true);
        return false;
    }
    se.items.borrow_mut().extend(added);
    true
}

/// 후보 목록을 직접 그린다. sel 행이 보이도록 top(스크롤 오프셋)도 여기서
/// 보정한다 — 행 높이는 폰트 실측이 필요해 paint 시점에만 정확하다.
fn paint(se: &Session, hwnd: &w::HWND) -> w::SysResult<()> {
    let hdc = hwnd.BeginPaint()?;
    let font = w::HFONT::GetStockObject(co::STOCK_FONT::DEFAULT_GUI)?;
    let _old_font = hdc.SelectObject(&font)?;

    let tm = hdc.GetTextMetrics()?;
    let row_h = tm.tmHeight + 2;
    se.row_h.set(row_h);

    let rc = hwnd.GetClientRect()?;
    let visible = (((rc.bottom - rc.top) / row_h).max(1)) as usize;

    let items = se.items.borrow();
    let sel = se.sel.get();
    let mut top = se.top.get();
    if sel < top {
        top = sel;
    } else if sel >= top + visible {
        top = sel + 1 - visible;
    }
    se.top.set(top);

    let end = items.len().min(top + visible);
    for (row, idx) in (top..end).enumerate() {
        let y = row as i32 * row_h;
        let (bg, fg) = if idx == sel {
            (co::COLOR::HIGHLIGHT, co::COLOR::HIGHLIGHTTEXT)
        } else {
            (co::COLOR::WINDOW, co::COLOR::WINDOWTEXT)
        };
        let rc_row = w::RECT { left: rc.left, top: y, right: rc.right, bottom: y + row_h };
        let brush = w::HBRUSH::GetSysColorBrush(bg)?;
        hdc.FillRect(rc_row, &brush)?;
        hdc.SetBkColor(w::GetSysColor(bg))?;
        hdc.SetTextColor(w::GetSysColor(fg))?;
        hdc.TextOut(2, y + 1, &items[idx])?;
    }
    Ok(())
}

/// 캐럿 화면 좌표 시도. 실패 시 화면 중앙.
///
/// 캐럿의 화면 y가 화면 높이의 2/3보다 더 낮은 위치(= 하단 1/3 영역)에 있으면
/// popup을 캐럿 위에 띄워 모니터 밖 이탈을 막는다. 그 외엔 캐럿 아래.
fn try_anchor() -> w::POINT {
    let screen_w = w::GetSystemMetrics(co::SM::CXSCREEN);
    let screen_h = w::GetSystemMetrics(co::SM::CYSCREEN);
    // popup 오른쪽이 화면 밖으로 잘리지 않도록 x 최대값 제한 (왼쪽은 0).
    let max_x = (screen_w - POPUP_WIDTH).max(0);
    let clamp_x = |x: i32| x.clamp(0, max_x);

    if let Some(fg) = w::HWND::GetForegroundWindow() {
        let (tid, _) = fg.GetWindowThreadProcessId();
        if tid != 0
            && let Ok(info) = w::GetGUIThreadInfo(tid)
        {
            let caret = &info.hwndCaret;
            if !caret.ptr().is_null()
                && (info.rcCaret.right - info.rcCaret.left) >= 0
                && (info.rcCaret.bottom - info.rcCaret.top) > 0
                && let Ok(top_pt) =
                    caret.ClientToScreen(w::POINT::with(info.rcCaret.left, info.rcCaret.top))
                && let Ok(bot_pt) =
                    caret.ClientToScreen(w::POINT::with(info.rcCaret.left, info.rcCaret.bottom))
            {
                let threshold = screen_h * 2 / 3;
                let caret_h = (bot_pt.y - top_pt.y).max(0);
                let margin = caret_h.max(40);
                if bot_pt.y > threshold {
                    // 캐럿이 화면 하단 1/3에 있음 → 위로. 음수로 빠지면
                    // 화면 상단으로 clamp.
                    let y = top_pt.y - POPUP_HEIGHT - margin;
                    return w::POINT::with(clamp_x(top_pt.x), y);
                } else {
                    return w::POINT::with(clamp_x(bot_pt.x), bot_pt.y + 2);
                }
            }
        }
    }
    w::POINT::with(clamp_x(screen_w / 2 - POPUP_WIDTH / 2), screen_h / 2 - POPUP_HEIGHT / 2)
}
