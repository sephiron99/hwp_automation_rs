use crate::com_util;
use crate::error::HwpError;
use crate::variant::FromVariant;
use std::mem::ManuallyDrop;
use windows::Win32::System::Com::IDispatch;
use windows::Win32::System::Variant::*;

/// C++ COleDispatchDriver 래퍼 클래스에 대응하는 Rust newtype을 생성합니다.
///
/// 생성되는 구현:
/// - `Deref<Target = DispObj>` — 범용 `get`/`call`/`call_with` 접근
/// - `FromVariant` — COM 메서드 반환값에서 자동 변환
/// - `IntoVariant` — COM 메서드 인자로 전달 가능
macro_rules! hwp_com_type {
    ($(#[$meta:meta])* $name:ident) => {
        $(#[$meta])*
        pub struct $name(crate::disp_obj::DispObj);

        impl std::ops::Deref for $name {
            type Target = crate::disp_obj::DispObj;
            fn deref(&self) -> &Self::Target {
                &self.0
            }
        }

        impl crate::variant::FromVariant for $name {
            fn from_variant(
                v: &windows::Win32::System::Variant::VARIANT,
            ) -> crate::error::Result<Self> {
                Ok($name(crate::variant::FromVariant::from_variant(v)?))
            }
        }

        impl crate::variant::IntoVariant for $name {
            fn into_variant(
                self,
            ) -> crate::error::Result<windows::Win32::System::Variant::VARIANT> {
                self.0.into_variant()
            }
        }
    };
}

/// 임의의 IDispatch COM 객체를 감싸는 경량 래퍼.
///
/// HWP의 sub-object (XHwpWindows, XHwpToolbarLayout 등)를 체이닝할 때 사용합니다.
///
/// ```ignore
/// let windows: DispObj = hwp.get("XHwpWindows")?;
/// let window: DispObj = windows.get("Active_XHwpWindow")?;
/// let layout: DispObj = window.get("XHwpToolbarLayout")?;
/// ```
#[derive(Clone)]
pub struct DispObj(pub(crate) IDispatch);

impl DispObj {
    pub fn get<T: FromVariant>(&self, name: &str) -> crate::error::Result<T> {
        let var = com_util::get_property(&self.0, name)?;
        T::from_variant(&var)
    }

    pub fn call<T: FromVariant>(&self, name: &str) -> crate::error::Result<T> {
        let var = com_util::call_method(&self.0, name)?;
        T::from_variant(&var)
    }

    pub fn call_with<T: FromVariant>(
        &self,
        name: &str,
        args: Vec<VARIANT>,
    ) -> crate::error::Result<T> {
        let var = com_util::call_method_with(&self.0, name, args)?;
        T::from_variant(&var)
    }

    /// 인자 있는 프로퍼티 읽기 (예: `Item(index)`)
    pub fn get_with<T: FromVariant>(
        &self,
        name: &str,
        args: Vec<VARIANT>,
    ) -> crate::error::Result<T> {
        let var = com_util::get_property_with(&self.0, name, args)?;
        T::from_variant(&var)
    }

    /// 프로퍼티 쓰기
    pub fn put<V: crate::variant::IntoVariant>(
        &self,
        name: &str,
        value: V,
    ) -> crate::error::Result<()> {
        com_util::put_property(&self.0, name, value.into_variant()?)
    }

    /// IDispatch를 소비하지 않고 VARIANT로 변환 (COM 메서드 인자용)
    pub fn to_variant(&self) -> crate::error::Result<VARIANT> {
        let mut v = VARIANT::default();
        unsafe {
            let inner = &mut v.Anonymous.Anonymous;
            inner.vt = VT_DISPATCH;
            inner.Anonymous.pdispVal = ManuallyDrop::new(Some(self.0.clone()));
        }
        Ok(v)
    }
}

impl FromVariant for DispObj {
    fn from_variant(v: &VARIANT) -> crate::error::Result<Self> {
        unsafe {
            let vt = v.Anonymous.Anonymous.vt;
            match vt {
                VT_DISPATCH => {
                    let pdispval = &v.Anonymous.Anonymous.Anonymous.pdispVal;
                    let dispatch = (**pdispval)
                        .clone()
                        .ok_or_else(|| HwpError::VariantConversion("null IDispatch".into()))?;
                    Ok(DispObj(dispatch))
                }
                _ => Err(HwpError::VariantConversion(format!(
                    "IDispatch 변환 불가 (VT: {})",
                    vt.0
                ))),
            }
        }
    }
}

impl crate::variant::IntoVariant for DispObj {
    fn into_variant(self) -> crate::error::Result<VARIANT> {
        let mut v = VARIANT::default();
        unsafe {
            let inner = &mut v.Anonymous.Anonymous;
            inner.vt = VT_DISPATCH;
            inner.Anonymous.pdispVal = ManuallyDrop::new(Some(self.0));
        }
        Ok(v)
    }
}
