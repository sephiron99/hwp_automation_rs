use crate::error::HwpError;
use std::mem::ManuallyDrop;
use windows::core::BSTR;
use windows::Win32::Foundation::VARIANT_BOOL;
use windows::Win32::System::Variant::*;

/// COM VARIANT → Rust 타입 변환
pub trait FromVariant: Sized {
    fn from_variant(v: &VARIANT) -> crate::error::Result<Self>;
}

/// Rust 타입 → COM VARIANT 변환
pub trait IntoVariant {
    fn into_variant(self) -> crate::error::Result<VARIANT>;
}

// =========================================================================
// VARIANT 내부 필드 접근 헬퍼
// =========================================================================

unsafe fn variant_inner(v: &mut VARIANT) -> &mut VARIANT_0_0 {
    unsafe { &mut v.Anonymous.Anonymous }
}

// =========================================================================
// FromVariant 구현
// =========================================================================

impl FromVariant for i32 {
    fn from_variant(v: &VARIANT) -> crate::error::Result<Self> {
        unsafe {
            let vt = v.Anonymous.Anonymous.vt;
            match vt {
                VT_I4 | VT_INT => Ok(v.Anonymous.Anonymous.Anonymous.lVal),
                _ => Err(HwpError::VariantConversion(format!(
                    "i32 변환 불가 (VT: {})",
                    vt.0
                ))),
            }
        }
    }
}

impl FromVariant for u32 {
    fn from_variant(v: &VARIANT) -> crate::error::Result<Self> {
        unsafe {
            let vt = v.Anonymous.Anonymous.vt;
            match vt {
                VT_UI4 => Ok(v.Anonymous.Anonymous.Anonymous.ulVal),
                VT_I4 | VT_INT => Ok(v.Anonymous.Anonymous.Anonymous.lVal as u32),
                _ => Err(HwpError::VariantConversion(format!(
                    "u32 변환 불가 (VT: {})",
                    vt.0
                ))),
            }
        }
    }
}

impl FromVariant for bool {
    fn from_variant(v: &VARIANT) -> crate::error::Result<Self> {
        unsafe {
            let vt = v.Anonymous.Anonymous.vt;
            match vt {
                VT_BOOL => Ok(v.Anonymous.Anonymous.Anonymous.boolVal.0 != 0),
                _ => Err(HwpError::VariantConversion(format!(
                    "bool 변환 불가 (VT: {})",
                    vt.0
                ))),
            }
        }
    }
}

impl FromVariant for f64 {
    fn from_variant(v: &VARIANT) -> crate::error::Result<Self> {
        unsafe {
            let vt = v.Anonymous.Anonymous.vt;
            match vt {
                VT_R8 => Ok(v.Anonymous.Anonymous.Anonymous.dblVal),
                _ => Err(HwpError::VariantConversion(format!(
                    "f64 변환 불가 (VT: {})",
                    vt.0
                ))),
            }
        }
    }
}

impl FromVariant for String {
    fn from_variant(v: &VARIANT) -> crate::error::Result<Self> {
        unsafe {
            let vt = v.Anonymous.Anonymous.vt;
            match vt {
                VT_BSTR => {
                    let bstr = &v.Anonymous.Anonymous.Anonymous.bstrVal;
                    Ok(bstr.to_string())
                }
                _ => Err(HwpError::VariantConversion(format!(
                    "String 변환 불가 (VT: {})",
                    vt.0
                ))),
            }
        }
    }
}

impl FromVariant for () {
    fn from_variant(_v: &VARIANT) -> crate::error::Result<Self> {
        Ok(())
    }
}

// =========================================================================
// IntoVariant 구현
// =========================================================================

impl IntoVariant for i32 {
    fn into_variant(self) -> crate::error::Result<VARIANT> {
        let mut v = VARIANT::default();
        unsafe {
            let inner = variant_inner(&mut v);
            inner.vt = VT_I4;
            inner.Anonymous.lVal = self;
        }
        Ok(v)
    }
}

impl IntoVariant for u32 {
    fn into_variant(self) -> crate::error::Result<VARIANT> {
        let mut v = VARIANT::default();
        unsafe {
            let inner = variant_inner(&mut v);
            inner.vt = VT_UI4;
            inner.Anonymous.ulVal = self;
        }
        Ok(v)
    }
}

impl IntoVariant for bool {
    fn into_variant(self) -> crate::error::Result<VARIANT> {
        let mut v = VARIANT::default();
        unsafe {
            let inner = variant_inner(&mut v);
            inner.vt = VT_BOOL;
            inner.Anonymous.boolVal = VARIANT_BOOL(if self { -1 } else { 0 });
        }
        Ok(v)
    }
}

impl IntoVariant for f64 {
    fn into_variant(self) -> crate::error::Result<VARIANT> {
        let mut v = VARIANT::default();
        unsafe {
            let inner = variant_inner(&mut v);
            inner.vt = VT_R8;
            inner.Anonymous.dblVal = self;
        }
        Ok(v)
    }
}

impl IntoVariant for &str {
    fn into_variant(self) -> crate::error::Result<VARIANT> {
        let bstr = BSTR::from(self);
        let mut v = VARIANT::default();
        unsafe {
            let inner = variant_inner(&mut v);
            inner.vt = VT_BSTR;
            inner.Anonymous.bstrVal = ManuallyDrop::new(bstr);
        }
        Ok(v)
    }
}

impl IntoVariant for String {
    fn into_variant(self) -> crate::error::Result<VARIANT> {
        self.as_str().into_variant()
    }
}
