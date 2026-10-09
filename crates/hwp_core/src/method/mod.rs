//! SDK method 층 — `HwpObject.tlb.IDL` interface의 method/property 1:1 binding
//!
//! 호출 하나가 COM 호출 하나에 대응한다. 여러 호출을 묶거나 기본값을 정하는
//! 코드는 [`crate::ext`]에 둔다.
//!
//! - [`hwp_object`] — `IHwpObject`
//! - [`movepos`] — `IHwpObject::MovePos`
//! - [`h_action`] — `HAction`
//! - [`xhwp`] — `IXHwp*` 하위 객체
pub mod h_action;
pub mod hwp_object;
pub mod movepos;
pub mod xhwp;
