# hwp_addon 크레이트 분석

HWP가 로드하는 Add-in DLL을 Rust로 작성하기 위한 프레임워크.

---

## 개요

HWP는 `QueryUserActionInterface` 함수를 export하는 DLL을 애드인으로 로드한다.
`hwp_addon`은 C++ `IHncUserActionModule` vtable 호환 레이아웃을 Rust로 구현하고,
개발자가 `HwpUserAction` 트레잇만 구현하면 DLL을 완성할 수 있도록 래핑한다.

---

## 모듈 구조

```
hwp_addon/src/
├── lib.rs              # export_hwp_addon! 매크로
├── ffi.rs              # C++ vtable FFI 정의 + 트램펄린 함수
├── hwp_user_action.rs  # HwpUserAction 트레잇, ToolbarConfig 등 공개 API
├── toolbar.rs          # ToolbarBitmap — BMP 로드 및 HBITMAP 캐시
└── debug.rs            # 파일 로그 + MessageBox 유틸
```

---

## 구조도

### 레이어 구성

```
┌──────────────────────────────────────────────────────────────────┐
│                          hwp_addon                               │
│                                                                  │
│  ┌────────────────────────────────────────────────────────────┐  │
│  │  lib.rs  ─  export_hwp_addon! 매크로                       │  │
│  │                                                            │  │
│  │  생성물:                                                   │  │
│  │    static VTABLE : IHncUserActionModuleVtbl                │  │
│  │    static MODULE : RustActionModule<T>                     │  │
│  │    pub fn QueryUserActionInterface() → *const MODULE       │  │
│  │    pub fn IsAccessiblePath(…) → i32                        │  │
│  └──────────────────────────┬─────────────────────────────────┘  │
│                             │ 참조                               │
│  ┌──────────────────────────▼─────────────────────────────────┐  │
│  │  ffi.rs  ─  C++ vtable 호환 레이어                         │  │
│  │                                                            │  │
│  │  IHncUserActionModuleVtbl (repr C)                         │  │
│  │    EnumAction      : extern "thiscall" fn ptr              │  │
│  │    GetActionImage  : extern "thiscall" fn ptr              │  │
│  │    UpdateUI        : extern "thiscall" fn ptr              │  │
│  │    DoAction        : extern "thiscall" fn ptr              │  │
│  │                                                            │  │
│  │  RustActionModule<T: HwpUserAction> (repr C)               │  │
│  │    lpVtbl : *const IHncUserActionModuleVtbl   ← [0..4]     │  │
│  │    plugin : T                                 ← [4..]      │  │
│  │                                                            │  │
│  │  트램펄린 (thiscall → trait 위임)                          │  │
│  │    tramp_enum_action       → plugin.enum_action()          │  │
│  │    tramp_get_action_image  → plugin.get_action_image()     │  │
│  │    tramp_update_ui         → plugin.dispatch_update_ui()   │  │
│  │    tramp_do_action         → plugin.dispatch()             │  │
│  └──────────────────────────┬─────────────────────────────────┘  │
│                             │ 위임                               │
│  ┌──────────────────────────▼─────────────────────────────────┐  │
│  │  hwp_user_action.rs  ─  트레잇 레이어                      │  │
│  │                                                            │  │
│  │  trait HwpUserAction                                       │  │
│  │    ┌─ 필수 구현 ──────────────────────────────────────┐    │  │
│  │    │  fn actions()   → &[ActionMeta]                  │    │  │
│  │    │  fn do_action() → Result<bool>                   │    │  │
│  │    └──────────────────────────────────────────────────┘    │  │
│  │    ┌─ 선택 구현 (기본 제공) ──────────────────────────┐    │  │
│  │    │  fn toolbar_config()  → Option<&ToolbarConfig>   │    │  │
│  │    │  fn on_initial_load() → Result<bool>             │    │  │
│  │    │  fn on_load()         → Result<bool>             │    │  │
│  │    │  fn update_ui()       → u32                      │    │  │
│  │    │  fn toolbar_bitmap()  → Option<isize>            │    │  │
│  │    └──────────────────────────────────────────────────┘    │  │
│  │    ┌─ 프레임워크 내부 (오버라이드 불필요) ────────────┐    │  │
│  │    │  fn enum_action()          → Option<&str>        │    │  │
│  │    │  fn get_action_image()     → Option<(isize,i32)> │    │  │
│  │    │  fn dispatch()             → Result<bool>        │    │  │
│  │    │  fn dispatch_update_ui()   → u32                 │    │  │
│  │    │  fn setup_toolbar()        → Result<bool>        │    │  │
│  │    └──────────────────────────────────────────────────┘    │  │
│  │                                                            │  │
│  │  ActionMeta  { name, label, image_index }                  │  │
│  │  ToolbarConfig { name, serialize_path, bitmap_data,        │  │
│  │                  target, ribbon_tab, ribbon_toolbox_index }│  │
│  │  ToolbarTarget { Toolbar | Ribbon | Both }                 │  │
│  │  RibbonTab     { New(&str) }                               │  │
│  └──────────────────────────┬─────────────────────────────────┘  │
│                             │ 사용                               │
│  ┌──────────────────────────▼─────────────────────────────────┐  │
│  │  toolbar.rs  ─  GDI 비트맵 캐시                            │  │
│  │                                                            │  │
│  │  ToolbarBitmap { inner: OnceLock<Option<isize>> }          │  │
│  │    fn load(path)          — LoadImageW                     │  │
│  │    fn load_from_bytes(data) — CreateDIBitmap               │  │
│  │    fn image(index) → Option<(isize, i32)>                  │  │
│  └────────────────────────────────────────────────────────────┘  │
│                                                                  │
│  ┌────────────────────────────────────────────────────────────┐  │
│  │  debug.rs  ─  로그/MessageBox 유틸                         │  │
│  │    log(tag, msg) / msgbox(title, msg) / msgbox_err(…)      │  │
│  └────────────────────────────────────────────────────────────┘  │
└──────────────────────────────────────────────────────────────────┘
```

### 호출 흐름

```
HWP 프로세스                          hwp_addon DLL
────────────────────────────────────────────────────────────────
LoadLibrary("addon.dll")
  │
  ├─→ QueryUserActionInterface()
  │       └─→ &MODULE  (RustActionModule<T>, static)
  │                ├── lpVtbl → VTABLE (IHncUserActionModuleVtbl)
  │                └── plugin : T (HwpUserAction 구현체)
  │
  ├─→ EnumAction(n)
  │       └─→ tramp_enum_action
  │               └─→ extract_plugin(this) → plugin
  │                       └─→ plugin.enum_action(n)
  │                             0 → UUIDSTR_ON_INITIAL_LOAD
  │                             1 → UUIDSTR_ON_LOAD
  │                             n → actions()[n-2].name
  │
  ├─→ GetActionImage(name, state, phBitmap, pnIndex)
  │       └─→ tramp_get_action_image
  │               └─→ plugin.get_action_image()
  │                       └─→ plugin.toolbar_bitmap(state)
  │                               └─→ ToolbarBitmap::image()
  │
  ├─→ UpdateUI(name, pDisp, lpuState)
  │       └─→ tramp_update_ui
  │               └─→ with_hwp_object(pDisp) → HwpObject (mem::forget)
  │                       └─→ plugin.dispatch_update_ui(name, hwp)
  │                             예약 GUID → 0
  │                             그 외    → plugin.update_ui()
  │
  └─→ DoAction(name, pDisp)
          └─→ tramp_do_action
                  └─→ with_hwp_object(pDisp) → HwpObject (mem::forget)
                          └─→ plugin.dispatch(name, hwp)
                                ON_INITIAL_LOAD → on_initial_load()
                                ON_LOAD         → on_load()
                                                    └─→ setup_toolbar()
                                그 외            → do_action()
```

---

## FFI 레이어 (`ffi.rs`)

### C++ 인터페이스 대응

C++ 원본:
```cpp
interface IHncUserActionModule {
    virtual LPCSTR EnumAction(int) = 0;
    virtual BOOL   GetActionImage(LPCSTR, UINT, HBITMAP*, int*) = 0;
    virtual BOOL   UpdateUI(LPCSTR, LPDISPATCH, UINT*) = 0;
    virtual int    DoAction(LPCSTR, LPDISPATCH) = 0;
};
```

MSVC 32-bit에서 가상 함수는 `__thiscall` 호출 규약(this → ECX)을 사용한다.
`IHncUserActionModuleVtbl`은 이를 그대로 `extern "thiscall"` 함수 포인터로 구성한다.

### 메모리 레이아웃

```
RustActionModule<T>
 ├── [0..4]  lpVtbl  →  IHncUserActionModuleVtbl  (C++ vptr 위치)
 └── [4..]   plugin  :  T  (HwpUserAction 구현체)
```

HWP가 vtable 포인터 바로 뒤에 `this`로 접근하므로 `#[repr(C)]` 필수.

### 트램펄린 함수

`tramp_enum_action`, `tramp_get_action_image`, `tramp_update_ui`, `tramp_do_action`
네 함수가 C++ vtable 호출을 받아 Rust 트레잇 메서드로 위임한다.

**IDispatch 수명 처리**:
HWP가 소유한 `IDispatch` 포인터를 `HwpObject`로 래핑할 때
`AddRef`/`Release`를 호출하면 안 된다.
`with_hwp_object`가 `from_raw_dispatch` 후 `std::mem::forget`으로 Drop을 방지한다.

**EnumAction 문자열 버퍼**:
반환하는 `*const c_char`는 `thread_local!` `Vec<u8>`에 보관한다.
HWP는 다음 `EnumAction` 호출 전에 문자열을 복사하므로 안전하다.

---

## 플러그인 트레잇 (`hwp_user_action.rs`)

### `HwpUserAction`

개발자가 구현하는 트레잇. 필수/선택 구분:

| 메서드              | 필수 | 설명                                                 |
|---------------------|------|------------------------------------------------------|
| `actions()`         | ✓    | 사용자 액션 목록 반환                                |
| `do_action()`       | ✓    | 액션 실행                                            |
| `toolbar_config()`  | —    | 툴바/리본 자동 설정용 설정 반환                      |
| `on_initial_load()` | —    | 최초 등록 시 한 번 호출                              |
| `on_load()`         | —    | HWP 창 열릴 때마다 호출 (기본: `setup_toolbar` 실행) |
| `update_ui()`       | —    | 버튼 UI 상태 반환                                    |
| `toolbar_bitmap()`  | —    | HBITMAP 반환 (toolbar_config 구현 시 자동)           |

    ### 액션 열거 순서

`enum_action(n)`이 반환하는 순서:

```
0 → UUIDSTR_ON_INITIAL_LOAD  ({B91A2981-...})
1 → UUIDSTR_ON_LOAD          ({3E4DC866-...})
2 → actions()[0].name
3 → actions()[1].name
...
```

예약 GUID는 프레임워크가 자동 처리하므로 `actions()`에 넣지 않는다.

### `dispatch` / `dispatch_update_ui`

예약 GUID를 라이프사이클 훅으로, 나머지를 `do_action`으로 라우팅한다.

---

## 툴바/리본 자동 설정

### `ToolbarConfig`

```rust
pub struct ToolbarConfig {
    pub name: &'static str,          // 툴바 이름
    pub serialize_path: &'static str, // HWP 직렬화 경로 (리본 탭 UID로도 사용)
    pub bitmap_data: &'static [u8],  // 960×16 32bit RGBA BMP (include_bytes!)
    pub target: ToolbarTarget,       // Toolbar / Ribbon / Both
    pub ribbon_toolbox_index: i32,   // -1: 끝에 추가
}
```

### `setup_toolbar` 동작

1. `ChangeSerializePath`로 serialize_path 설정
2. `IsNewSerializePath`가 false면 이미 등록된 것이므로 바로 반환
3. 툴바 생성: `CreateToolbar` → 각 액션마다 `CreateToolbarButton` → `InsertButton`
4. 리본 생성: `GetToolboxToolbar` → `InsertToolboxTab` → `InsertToolbox` → `GetLayout` → `InsertGroup` → 각 액션마다 `CreateToolboxItemButtonEx` → `InsertItem`

### `ToolbarBitmap`

`OnceLock<Option<isize>>`으로 HBITMAP을 한 번만 로드한다.
`load_from_bytes`는 BMP 바이너리에서 직접 `CreateDIBitmap`으로 GDI 비트맵을 생성한다.
파일 경로 로드는 `load` 메서드 (`LoadImageW`).

---

## `export_hwp_addon!` 매크로 (`lib.rs`)

플러그인 타입을 받아 두 개의 `no_mangle` export 함수를 생성한다:

```rust
export_hwp_addon!(MyPlugin, MyPlugin);
```

생성물:
- **`QueryUserActionInterface`**: `&MODULE`(static `RustActionModule`) 포인터 반환
- **`IsAccessiblePath`**: 항상 1(TRUE) 반환

`VTABLE`(static `IHncUserActionModuleVtbl`)과 `MODULE`(static `RustActionModule<T>`)도 생성한다.
둘 다 `'static`이므로 DLL 수명 동안 유효하다.

---

## 디버그 유틸 (`debug.rs`)

| 함수                     | 동작                                                                  |
|--------------------------|-----------------------------------------------------------------------|
| `log(tag, msg)`          | `%LOCALAPPDATA%\hwp_addon_debug.log`에 타임스탬프 포함 기록 |
| `msgbox(title, msg)`     | 로그 + 정보 MessageBox                                                |
| `msgbox_err(title, msg)` | 로그 + 에러 MessageBox                                                |

---

## 사용 예시 (`hello_hwp_dll`)

```rust
pub struct HelloWorldPlugin;

impl HwpUserAction for HelloWorldPlugin {
    fn toolbar_config(&self) -> Option<&ToolbarConfig> { Some(&CONFIG) }

    fn actions(&self) -> &[ActionMeta] {
        &[ActionMeta { name: "InsertHelloWorld", label: "Hello", image_index: 0 }]
    }

    fn do_action(&self, action_name: &str, hwp: &HwpObject) -> Result<bool> {
        // hwp.h_action(), hwp.h_parameter_set() 등으로 HWP 제어
        Ok(true)
    }
}

export_hwp_addon!(HelloWorldPlugin, HelloWorldPlugin);
```

---

## 의존성

- **`hwp_core`**: `HwpObject`, `HwpVer`, COM 유틸
- **`windows`**: Win32 API (GDI, UI, COM, Registry 등)
- **`widestring`**: UTF-16 문자열 변환
- **`thiserror`**: 에러 타입
- **`chrono`**: 로그 타임스탬프
