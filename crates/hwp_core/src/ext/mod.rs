//! 편의 API 층 — SDK method와 Action을 조합한 고수준 함수
//!
//! SDK에 1:1로 대응하지 않는, 이 프로젝트가 자체 설계한 API다. 아래 층
//! ([`crate::pset`], [`crate::method`], [`crate::actions`])의 공개 API만으로 구현한다.
//!
//! SDK binding(`hwp.get_text()`, `hwp.h_action()?.undo()`)과 호출 자리에서 구분되도록,
//! 편의 API는 모두 `ext_` 접두어가 붙은 핸들을 거쳐 부른다. import는 필요 없다.
//!
//! ```ignore
//! hwp.ext_text().insert("Hello")?;
//! let shape = hwp.ext_shape().current_char()?;
//! hwp.ext_doc().save_as_hwpx("C:\\out.hwpx")?;
//! ```
//!
//! - [`HwpObject::ext_text`] → [`ExtText`] — 텍스트 삽입, 문서 텍스트 순회
//! - [`HwpObject::ext_shape`] → [`ExtShape`] — 커서 위치의 글자·문단 모양
//! - [`HwpObject::ext_doc`] → [`ExtDoc`] — HWPX 저장, 커서 위치가 표시된 HWPML·HWPX
//!
//! [`HwpObject::ext_text`]: crate::HwpObject::ext_text
//! [`HwpObject::ext_shape`]: crate::HwpObject::ext_shape
//! [`HwpObject::ext_doc`]: crate::HwpObject::ext_doc
pub mod char_shape;
pub mod doc;
pub mod hwpml_caret;
pub mod para_shape;
pub mod shape;
pub mod text;

pub use char_shape::CharShape;
pub use doc::ExtDoc;
pub use hwpml_caret::{HwpmlCaret, HwpxCaret};
pub use para_shape::ParaShape;
pub use shape::ExtShape;
pub use text::{ExtText, TextSegment, TextWalker};
