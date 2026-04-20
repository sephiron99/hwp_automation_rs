use crate::error::Result;
use std::mem::ManuallyDrop;
use windows::core::{BSTR, GUID, PCWSTR};
use windows::Win32::Globalization::LOCALE_USER_DEFAULT;
use windows::Win32::System::Com::{
    IDispatch, DISPATCH_FLAGS, DISPATCH_METHOD, DISPATCH_PROPERTYGET, DISPPARAMS, EXCEPINFO,
};
use windows::Win32::System::Variant::*;

/// IDispatch::Invoke를 호출하는 내부 함수
fn invoke(
    dispatch: &IDispatch,
    name: &str,
    flags: DISPATCH_FLAGS,
    args: &mut [VARIANT],
) -> Result<VARIANT> {
    let kind = if flags == DISPATCH_PROPERTYGET {
        "GET"
    } else {
        "CALL"
    };
    // crate::debug::log("com_util", &format!("{kind} {name} (args={})", args.len()));

    let bstr_name = BSTR::from(name);
    let mut dispid: i32 = 0;

    unsafe {
        dispatch.GetIDsOfNames(
            &GUID::zeroed(),
            &PCWSTR(bstr_name.as_ptr()),
            1,
            LOCALE_USER_DEFAULT,
            &mut dispid,
        )?;
    }

    // crate::debug::log("com_util", &format!("{kind} {name} → dispid={dispid}"));

    let dp = DISPPARAMS {
        cArgs: args.len() as u32,
        rgvarg: if args.is_empty() {
            std::ptr::null_mut()
        } else {
            args.as_mut_ptr()
        },
        cNamedArgs: 0,
        rgdispidNamedArgs: std::ptr::null_mut(),
    };

    let mut result = VARIANT::default();
    let mut excep_info = EXCEPINFO::default();
    let mut arg_err: u32 = 0;

    let hr = unsafe {
        dispatch.Invoke(
            dispid,
            &GUID::zeroed(),
            LOCALE_USER_DEFAULT,
            flags,
            &dp,
            Some(&mut result),
            Some(&mut excep_info),
            Some(&mut arg_err),
        )
    };

    if let Err(e) = hr {
        let desc = if !excep_info.bstrDescription.is_empty() {
            Some(excep_info.bstrDescription.to_string())
        } else if !excep_info.bstrSource.is_empty() {
            Some(excep_info.bstrSource.to_string())
        } else {
            None
        };

        if let Some(desc) = desc {
            crate::debug::log("com_util", &format!("{kind} {name} → EXCEP: {desc}"));
            return Err(crate::error::HwpError::ExecutionFailed(format!(
                "{name}: {desc}"
            )));
        }

        crate::debug::log("com_util", &format!("{kind} {name} → ERR: {e}"));
        return Err(e.into());
    }

    // crate::debug::log("com_util", &format!("{kind} {name} → OK"));
    Ok(result)
}

/// 프로퍼티 값을 읽습니다.
pub fn get_property(dispatch: &IDispatch, name: &str) -> Result<VARIANT> {
    invoke(dispatch, name, DISPATCH_PROPERTYGET, &mut [])
}

/// 인자 있는 프로퍼티 값을 읽습니다. (예: `Item(index)`)
pub fn get_property_with(dispatch: &IDispatch, name: &str, args: Vec<VARIANT>) -> Result<VARIANT> {
    let mut reversed: Vec<VARIANT> = args.into_iter().rev().collect();
    invoke(dispatch, name, DISPATCH_PROPERTYGET, &mut reversed)
}

/// 인자 없는 메서드를 호출합니다.
pub fn call_method(dispatch: &IDispatch, name: &str) -> Result<VARIANT> {
    invoke(dispatch, name, DISPATCH_METHOD, &mut [])
}

/// 인자 있는 메서드를 호출합니다. (인자는 COM 규약에 따라 역순 배치됩니다)
pub fn call_method_with(dispatch: &IDispatch, name: &str, args: Vec<VARIANT>) -> Result<VARIANT> {
    let mut reversed: Vec<VARIANT> = args.into_iter().rev().collect();
    invoke(dispatch, name, DISPATCH_METHOD, &mut reversed)
}

/// BSTR* 출력 매개변수가 있는 메서드를 호출합니다. (예: GetText)
/// 반환값: (메서드 반환값 VARIANT, 출력 BSTR 문자열)
pub fn call_method_with_bstr_out(
    dispatch: &IDispatch,
    name: &str,
    extra_args: Vec<VARIANT>,
) -> Result<(VARIANT, String)> {
    let bstr_name = BSTR::from(name);
    let mut dispid: i32 = 0;

    unsafe {
        dispatch.GetIDsOfNames(
            &GUID::zeroed(),
            &PCWSTR(bstr_name.as_ptr()),
            1,
            LOCALE_USER_DEFAULT,
            &mut dispid,
        )?;
    }

    // BSTR* 출력 매개변수 준비
    let mut out_bstr = ManuallyDrop::new(BSTR::new());
    let mut bstr_var = VARIANT::default();
    unsafe {
        let inner = &mut bstr_var.Anonymous.Anonymous;
        inner.vt = VARENUM(VT_BSTR.0 | VT_BYREF.0);
        inner.Anonymous.pbstrVal = &mut out_bstr as *mut ManuallyDrop<BSTR> as *mut BSTR;
    }

    // 인자 배열 구성 (COM 규약: 역순, BSTR*가 마지막 인자이므로 역순에서 첫 번째)
    let mut args: Vec<VARIANT> = extra_args.into_iter().rev().collect();
    args.insert(0, bstr_var);

    let dp = DISPPARAMS {
        cArgs: args.len() as u32,
        rgvarg: args.as_mut_ptr(),
        cNamedArgs: 0,
        rgdispidNamedArgs: std::ptr::null_mut(),
    };

    let mut result = VARIANT::default();
    let mut excep_info = EXCEPINFO::default();
    let mut arg_err: u32 = 0;

    let hr = unsafe {
        dispatch.Invoke(
            dispid,
            &GUID::zeroed(),
            LOCALE_USER_DEFAULT,
            DISPATCH_METHOD,
            &dp,
            Some(&mut result),
            Some(&mut excep_info),
            Some(&mut arg_err),
        )
    };

    if let Err(e) = hr {
        let desc = if !excep_info.bstrDescription.is_empty() {
            Some(excep_info.bstrDescription.to_string())
        } else if !excep_info.bstrSource.is_empty() {
            Some(excep_info.bstrSource.to_string())
        } else {
            None
        };

        if let Some(desc) = desc {
            crate::debug::log(
                "com_util",
                &format!("CALL(bstr_out) {name} → EXCEP: {desc}"),
            );
            return Err(crate::error::HwpError::ExecutionFailed(format!(
                "{name}: {desc}"
            )));
        }

        crate::debug::log("com_util", &format!("CALL(bstr_out) {name} → ERR: {e}"));
        return Err(e.into());
    }

    // crate::debug::log("com_util", &format!("CALL(bstr_out) {name} → OK"));

    let text = ManuallyDrop::into_inner(out_bstr).to_string();
    Ok((result, text))
}

/// 프로퍼티 값을 설정합니다. (DISPATCH_PROPERTYPUT)
pub fn put_property(dispatch: &IDispatch, name: &str, value: VARIANT) -> Result<()> {
    use windows::Win32::System::Com::DISPATCH_PROPERTYPUT;

    crate::debug::log("com_util", &format!("PUT {name}"));

    let bstr_name = BSTR::from(name);
    let mut dispid: i32 = 0;

    unsafe {
        dispatch.GetIDsOfNames(
            &GUID::zeroed(),
            &PCWSTR(bstr_name.as_ptr()),
            1,
            LOCALE_USER_DEFAULT,
            &mut dispid,
        )?;
    }

    crate::debug::log("com_util", &format!("PUT {name} → dispid={dispid}"));

    let mut args = [value];
    let mut named_arg: i32 = -3; // DISPID_PROPERTYPUT

    let dp = DISPPARAMS {
        cArgs: 1,
        rgvarg: args.as_mut_ptr(),
        cNamedArgs: 1,
        rgdispidNamedArgs: &mut named_arg,
    };

    let mut excep_info = EXCEPINFO::default();
    let mut arg_err: u32 = 0;

    let hr = unsafe {
        dispatch.Invoke(
            dispid,
            &GUID::zeroed(),
            LOCALE_USER_DEFAULT,
            DISPATCH_PROPERTYPUT,
            &dp,
            None,
            Some(&mut excep_info),
            Some(&mut arg_err),
        )
    };

    if let Err(e) = hr {
        let desc = if !excep_info.bstrDescription.is_empty() {
            Some(excep_info.bstrDescription.to_string())
        } else if !excep_info.bstrSource.is_empty() {
            Some(excep_info.bstrSource.to_string())
        } else {
            None
        };

        if let Some(desc) = desc {
            crate::debug::log("com_util", &format!("PUT {name} → EXCEP: {desc}"));
            return Err(crate::error::HwpError::ExecutionFailed(format!(
                "{name}: {desc}"
            )));
        }

        crate::debug::log("com_util", &format!("PUT {name} → ERR: {e}"));
        return Err(e.into());
    }

    // crate::debug::log("com_util", &format!("PUT {name} → OK"));
    Ok(())
}
