pub mod debug;
pub mod ffi;
pub mod hwp_user_action;
pub mod ime;
pub mod keyfwd;
pub mod shortcut;
pub mod text_edit;
pub mod text_extract;
pub mod toolbar;

/// HWP 사용자 액션 애드인 DLL에 필요한 정적 모듈과 export 진입점을 생성합니다.
///
/// 이 매크로는 [`HwpUserAction`] 구현체를 HWP의 C++ 사용자 액션 인터페이스와
/// 연결합니다. 플러그인 인스턴스와 함수 포인터 테이블을 정적 저장소에 배치하고,
/// HWP가 DLL을 로드할 때 찾는 두 개의 C ABI 함수를 export합니다.
///
/// # 인자
///
/// - `$plugin_type`: [`HwpUserAction`]을 구현한 플러그인 타입입니다. 생성되는
///   vtable의 콜백을 이 타입에 맞게 단형화하는 데 사용됩니다.
/// - `$plugin_instance`: `$plugin_type`의 정적 인스턴스를 만드는 표현식입니다.
///   단위 구조체는 타입 이름 자체를 전달할 수 있고, 상태가 있는 구조체는 모든
///   필드를 초기화한 생성식을 전달해야 합니다. 결과값은 DLL이 언로드될 때까지
///   유지되는 정적 모듈에 저장되므로 `const` 문맥에서 생성 가능해야 합니다.
///
/// # 생성되는 항목
///
/// - `VTABLE`: HWP의 `IHncUserActionModule`과 호환되는 함수 포인터 테이블입니다.
///   액션 열거, 아이콘 조회, UI 상태 갱신 및 액션 실행 콜백을 포함합니다.
/// - `MODULE`: vtable 포인터와 `$plugin_instance`를 함께 보관하는 정적 모듈입니다.
/// - `QueryUserActionInterface`: HWP가 DLL을 로드할 때 호출하는 export 함수입니다.
///   HWP는 반환된 모듈의 vtable을 통해 [`HwpUserAction::enum_action`],
///   [`HwpUserAction::get_action_image`], [`HwpUserAction::dispatch_update_ui`] 및
///   [`HwpUserAction::dispatch`]에 접근합니다.
/// - `IsAccessiblePath`: HWP가 파일 경로 접근 가능 여부를 조회할 때 호출하는
///   export 함수입니다. 현재 기본 구현은 모든 경로에 대해 `TRUE`를 반환합니다.
///
/// # 수명과 스레드 모델
///
/// 생성된 `VTABLE`과 `MODULE`은 정적 객체이므로 DLL이 로드된 동안 주소가 바뀌지
/// 않습니다. HWP가 보관하는 인터페이스 포인터는 이 정적 객체를 가리킵니다.
/// 플러그인의 가변 상태가 필요하면 HWP의 호출 방식과 재진입 가능성을 고려하여
/// [`RefCell`], [`Mutex`] 등의 내부 가변성 타입을 `$plugin_type` 안에서 사용해야
/// 합니다.
///
/// # 안전성
///
/// 생성되는 두 export 함수와 32비트 HWP의 C++ `thiscall` vtable 콜백은 원시
/// 포인터를 다루므로 내부적으로 `unsafe`입니다. 호출 규약 변환, COM `IDispatch`
/// 포인터의 임시 래핑 및 참조 카운트 보존은 프레임워크가 담당합니다. 플러그인
/// 구현자는 이 매크로를 정상적인 HWP 애드인 진입점으로 사용하는 한 직접 원시
/// 포인터를 처리할 필요가 없습니다.
///
/// 하나의 DLL에서 이 매크로를 두 번 이상 호출하면 동일한 export 함수와 정적 항목이
/// 중복 정의되므로, 애드인 크레이트당 정확히 한 번만 호출해야 합니다.
///
/// # 예제
///
/// ```ignore
/// use hwp_addon::export_hwp_addon;
/// use hwp_addon::hwp_user_action::HwpUserAction;
///
/// struct MyPlugin;
///
/// impl HwpUserAction for MyPlugin {
///     // actions, do_action 구현
/// }
///
/// export_hwp_addon!(MyPlugin, MyPlugin);
/// ```
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
        pub unsafe extern "system" fn QueryUserActionInterface()
        -> *const $crate::ffi::RustActionModule<$plugin_type> {
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
