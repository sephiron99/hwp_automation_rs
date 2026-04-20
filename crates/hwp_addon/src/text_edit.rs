//! addon에서 IME 걱정 없이 쓰는 고수준 텍스트 편집 헬퍼.
//!
//! # 왜 hwp_addon에 있는가
//!
//! `hwp_core`는 HWP COM/OLE의 1:1 바인딩 레이어다. 반면 한글 IME 조합 중에는
//! `Move*`/`Select*` 계열 [`HAction`]이 캐럿을 움직이지 못하는 문제가 있어,
//! addon이 단축키로 트리거되었을 때는 일반적인 액션 시퀀스만으로 텍스트
//! 편집이 깔끔히 되지 않는다.
//!
//! 이 문제는 [`crate::ime::commit_composition`]을 사용해 우회하는데, 이는
//! Win32 UI(TSF/SetFocus/IMM32)에 의존하므로 OLE 클라이언트 사용 사례인
//! `hwp_core`에 두기에는 layering이 어색하다. 그래서 addon framework 쪽인
//! 이곳에 헬퍼를 모은다.
//!
//! # IME 조합 계약
//!
//! 이 모듈의 메서드는 모두 **IME 조합이 이미 확정된 상태**를 가정한다.
//! [`HwpUserAction::dispatch`](crate::hwp_user_action::HwpUserAction::dispatch)와
//! [`shortcut::run_action`](crate::shortcut)이 addon `do_action` 진입 직전에
//! [`commit_composition`](crate::ime::commit_composition)을 호출하므로, 일반적인
//! addon 코드 (`do_action`/`on_load`/`update_ui`)에서 호출하면 항상 안전하다.
//!
//! 이 가정을 깨는 위치(예: 워커 스레드, 타이머 콜백 등)에서 호출해야 한다면
//! 직접 [`commit_composition`](crate::ime::commit_composition)을 먼저 부르면 된다.
//!
//! # 설계
//!
//! [`HwpObject`]에 메서드를 직접 추가하는 대신 extension trait
//! [`HwpEditExt`]로 노출한다. 이렇게 하면 `hwp_core`를 오염시키지 않으면서
//! addon 작성자는 `use hwp_addon::text_edit::HwpEditExt;` 한 줄로 모든
//! 헬퍼를 얻는다.
//!
//! # 사용 예
//! ```ignore
//! use hwp_addon::text_edit::HwpEditExt;
//!
//! fn do_action(&self, action_name: &str, hwp: &HwpObject) -> Result<bool> {
//!     // 커서 앞 prefix를 expansion으로 교체
//!     hwp.replace_chars_before(prefix.chars().count() as i32, &expansion)?;
//!     Ok(true)
//! }
//! ```
//!
//! [`HAction`]: hwp_core::h_action::HAction

use hwp_core::error::Result;
use hwp_core::hwp_obj::HwpObject;

/// [`HwpObject`]에 IME-safe 고수준 편집 메서드를 추가하는 extension trait.
///
/// 모듈 레벨 문서의 *IME 조합 계약* 절을 참고할 것.
pub trait HwpEditExt {
    /// 커서 바로 앞 `n`개 문자를 선택한다.
    ///
    /// - `n <= 0`이면 아무것도 하지 않고 `Ok(false)` 반환.
    /// - 문단 시작에 도달하면 가능한 만큼만 (즉 `[0..pos)`) 선택한다.
    /// - 선택할 게 없으면 (현재 pos가 0) `Ok(false)` 반환.
    /// - 성공적으로 선택하면 `Ok(true)`.
    ///
    /// 내부적으로 [`get_pos`](HwpObject::get_pos) + [`select_text`](HwpObject::select_text)를
    /// 사용한다. `Move*`/`Select*` 액션이 IME 조합 상태에서 동작하지 않는
    /// 문제를 우회하기 위함이다.
    fn select_chars_before(&self, n: i32) -> Result<bool>;

    /// 커서 앞 `prefix_chars` 문자를 `replacement`로 교체한다.
    ///
    /// 동작:
    /// 1. `prefix_chars > 0`이면 [`select_chars_before`](HwpEditExt::select_chars_before)로
    ///    prefix 영역 선택.
    /// 2. [`insert_text`](HwpObject::insert_text) 호출 — 선택 영역이 자동으로
    ///    `replacement`로 대체된다.
    /// 3. 캐럿은 `replacement`의 끝에 위치한다 (HWP의 `InsertText` 동작).
    ///
    /// `prefix_chars`가 0 이하이면 단순 [`insert_text`](HwpObject::insert_text)와
    /// 동일하다.
    ///
    /// dabbrev처럼 prefix 길이를 문자 수로 알고 있는 addon에 적합하다.
    /// prefix 문자열을 갖고 있다면 [`replace_word_before`](HwpEditExt::replace_word_before)가
    /// 더 편리하다.
    fn replace_chars_before(&self, prefix_chars: i32, replacement: &str) -> Result<()>;

    /// 커서 앞 `prefix` 길이만큼의 텍스트를 `replacement`로 교체한다.
    ///
    /// `prefix`의 *내용*이 그 위치에 정말 있는지는 검증하지 않는다 — 길이만
    /// 사용한다. addon 측에서 어차피 prefix 문자열을 직접 추출했으므로
    /// 일치 여부는 보장되어 있다는 가정.
    ///
    /// 내부적으로 [`replace_chars_before`](HwpEditExt::replace_chars_before)에
    /// `prefix.chars().count() as i32`를 넘긴다.
    fn replace_word_before(&self, prefix: &str, replacement: &str) -> Result<()>;
}

impl HwpEditExt for HwpObject {
    fn select_chars_before(&self, n: i32) -> Result<bool> {
        if n <= 0 {
            return Ok(false);
        }
        let (_, para, pos) = self.get_pos()?;
        let start = (pos - n).max(0);
        if start == pos {
            return Ok(false);
        }
        self.select_text(para, start, para, pos)
    }

    fn replace_chars_before(&self, prefix_chars: i32, replacement: &str) -> Result<()> {
        if prefix_chars > 0 {
            // 선택 실패(문단 시작 등)는 무시하고 그냥 삽입한다.
            let _ = self.select_chars_before(prefix_chars)?;
        }
        self.insert_text(replacement)
    }

    fn replace_word_before(&self, prefix: &str, replacement: &str) -> Result<()> {
        self.replace_chars_before(prefix.chars().count() as i32, replacement)
    }
}
