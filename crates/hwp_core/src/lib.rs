#![cfg_attr(docsrs, feature(doc_cfg))]

// 층 구조 (아래에서 위로만 의존): 기반 ← pset ← method ← actions ← ext
// 자세한 규칙은 저장소 루트의 hwp_core.md 참고.

// ── 기반 ──
pub mod com_util;
#[macro_use]
pub mod disp_obj;
pub mod error;
pub mod hwp_obj;
pub mod hwp_ver;
pub mod variant;

pub use hwp_obj::HwpObject;

// ── 1. ParameterSet 타입 ──
pub mod pset;
// ── 2. SDK method ──
pub mod method;
// ── 3. Action ──
pub mod actions;
// ── 4. 편의 API ──
pub mod ext;
