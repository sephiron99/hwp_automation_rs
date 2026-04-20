pub mod debug;
pub mod ffi;
pub mod hwp_user_action;
pub mod ime;
pub mod shortcut;
pub mod text_edit;
pub mod toolbar;

/// HWP 애드인 DLL의 필수 export 함수들을 생성합니다.
///
/// # 사용법
/// ```ignore
/// export_hwp_addon!(MyPlugin, MyPlugin);
/// ```
///
/// 이 매크로는 다음을 생성합니다:
/// - `QueryUserActionInterface` — HWP가 DLL 로드 시 호출하는 진입점
/// - `IsAccessiblePath` — 파일 접근 권한 검사 (기본: 항상 허용)
#[macro_export]
#[allow(non_snake_case)]
macro_rules! export_hwp_addon {
    ($plugin_type:ty, $plugin_instance:expr) => {
        static VTABLE: $crate::ffi::IHncUserActionModuleVtbl =
            $crate::ffi::IHncUserActionModuleVtbl {
                EnumAction: $crate::ffi::tramp_enum_action::<$plugin_type>,
                GetActionImage: $crate::ffi::tramp_get_action_image::<$plugin_type>,
                UpdateUI: $crate::ffi::tramp_update_ui::<$plugin_type>,
                DoAction: $crate::ffi::tramp_do_action::<$plugin_type>,
            };

        static MODULE: $crate::ffi::RustActionModule<$plugin_type> =
            $crate::ffi::RustActionModule {
                lpVtbl: &VTABLE,
                plugin: $plugin_instance,
            };

        #[unsafe(no_mangle)]
        pub unsafe extern "system" fn QueryUserActionInterface(
        ) -> *const $crate::ffi::RustActionModule<$plugin_type> {
            $crate::debug::log("hwp_addon", "QueryUserActionInterface 호출됨");
            &MODULE
        }

        #[unsafe(no_mangle)]
        pub unsafe extern "system" fn IsAccessiblePath(
            _hwnd: isize,
            _id: i32,
            _path: *const u16,
        ) -> i32 {
            1 // TRUE — 항상 허용
        }
    };
}
