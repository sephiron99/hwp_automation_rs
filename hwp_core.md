# hwp_core 구조도

## 크레이트 개요

Windows COM/OLE를 통해 한글(HWP)을 Rust에서 제어하는 공유 라이브러리. 모든 COM 호출, 타입 변환, 액션 메서드를 캡슐화하며 `hwp_com` 및 `hwp_addon`은 hwp_core 크레이트를 기반으로 한다.

---

## 모듈 트리

```
hwp_core (lib.rs)
│
├── error          - HwpError, Result<T>
├── hwp_ver        - HwpVer (버전 열거형)
├── variant        - FromVariant / IntoVariant 트레이트
├── com_util       - 저수준 COM Invoke 래퍼
├── disp_obj       - DispObj (일반 IDispatch 래퍼)
├── hwp_types      - hwp_com_type! 매크로 + COM 타입 구조체
├── hwp_obj        - HwpObject (최상위 API 진입점)
├── h_action       - HAction, HParameterSet, HInsertText
├── ihwp_object    - IHwpObject 메서드, 열거형, ScanRange
├── debug          - 디버그 로그 유틸
└── actions        - HwpObject에 대한 액션 메서드 모음
    ├── app        - 앱/파일/인쇄/프레젠테이션/버전
    ├── edit       - 편집/클립보드/선택/실행취소
    ├── find       - 찾기/바꾸기
    ├── chars      - 글자 모양
    ├── para       - 문단 모양/스타일/글머리표
    ├── move_      - 커서 이동 (139개)
    ├── insert     - 삽입
    ├── table      - 표 조작 (211개)
    ├── draw       - 그리기 개체/도형 (168개)
    ├── picture    - 그림 속성/효과
    ├── view       - 보기 옵션/화면 배율
    ├── page       - 편집 용지/바탕쪽/구역
    ├── note       - 각주/미주/주석
    ├── track      - 변경 추적
    └── macro_     - 매크로/빠른 교정/빠른 책갈피
```

---

## 핵심 타입

### `HwpObject` (`hwp_obj.rs`)

사용자에게 노출되는 최상위 API 진입점. `IHwpObject`에 대응.

```
HwpObject
├── dispatch: IDispatch       (pub(crate))
└── ver: HwpVer               (private)

Methods (hwp_obj.rs):
  new(dispatch: IDispatch) -> Result<Self>
  from_raw_dispatch(raw: *mut c_void) -> Result<Self>  [unsafe]
  version() -> &HwpVer
  get<T: FromVariant>(name: &str) -> Result<T>
  call<T: FromVariant>(name: &str) -> Result<T>
  call_with<T: FromVariant>(name: &str, args: Vec<VARIANT>) -> Result<T>

+ ihwp_object.rs impl:
  run(action: &str) -> Result<()>
  windows() -> Result<XHwpWindows>
  get_text_file(format, save_block) -> Result<String>
  init_scan(option, range, spara, spos, epara, epos) -> Result<bool>
  get_text() -> Result<(GetTextStatus, String)>
  release_scan() -> Result<()>

+ h_action.rs impl:
  h_action() -> Result<HAction>
  h_parameter_set() -> Result<HParameterSet>
  select() -> Result<()>
  move_word_begin() -> Result<()>
  insert_text(text: &str) -> Result<()>

+ actions/* impl:
  [600+ 액션 메서드들]
```

---

### `HwpVer` (`hwp_ver.rs`)

```
enum HwpVer
├── V2018(u8)       - major 10
├── V2022(u8)       - major 12
├── V2024(u8)       - major 13
└── Other(u8, u8)   - (major, minor)

Methods:
  from_u32(version_code: u32) -> Self
  from_version_string(v: &str) -> Self
  is_at_least(required: &HwpVer) -> bool
  display_name() -> String
  as_number() -> (u8, u8)
```

---

### `HwpError` (`error.rs`)

```
enum HwpError
├── ComError(windows::core::Error)
├── OleInitFailed
├── ConnectionFailed
├── ActionNotFound(String)
├── MissingParameter { action, param }
├── InvalidParameterType { param, details }
├── ExecutionFailed(String)
├── VariantConversion(String)
├── InvalidStringData
└── UnsupportedVersion { required_version, current_version }

type Result<T> = std::result::Result<T, HwpError>
```

---

### `DispObj` (`disp_obj.rs`)

일반 `IDispatch` 래퍼. `hwp_com_type!` 매크로로 생성되는 타입들의 기반.

```
struct DispObj(pub(crate) IDispatch)

Methods:
  get<T: FromVariant>(name: &str) -> Result<T>
  call<T: FromVariant>(name: &str) -> Result<T>
  call_with<T: FromVariant>(name: &str, args: Vec<VARIANT>) -> Result<T>
  get_with<T: FromVariant>(name: &str, args: Vec<VARIANT>) -> Result<T>
  put<V: IntoVariant>(name: &str, value: V) -> Result<()>
  to_variant() -> Result<VARIANT>

impl FromVariant, IntoVariant
```

---

### 변환 트레이트 (`variant.rs`)

```
trait FromVariant: Sized
  fn from_variant(v: &VARIANT) -> Result<Self>

trait IntoVariant
  fn into_variant(self) -> Result<VARIANT>

구현 타입:
  FromVariant: i32, u32, bool, f64, String, (), DispObj
  IntoVariant: i32, u32, bool, f64, &str, String, DispObj
```

---

### COM 유틸 (`com_util.rs`)

```
Public Functions:
  get_property(dispatch, name) -> Result<VARIANT>
  get_property_with(dispatch, name, args) -> Result<VARIANT>
  call_method(dispatch, name) -> Result<VARIANT>
  call_method_with(dispatch, name, args) -> Result<VARIANT>
  call_method_with_bstr_out(dispatch, name, extra_args) -> Result<(VARIANT, String)>
  put_property(dispatch, name, value) -> Result<()>

Private:
  invoke(dispatch, name, flags, args) -> Result<VARIANT>  [핵심 COM Invoke]
```

---

### COM 타입 구조체 (`hwp_types.rs`)

`hwp_com_type!` 매크로로 생성. 모두 `Deref<Target = DispObj>` 구현.

```
XHwpWindows        - HWP 창 컬렉션
  active() -> XHwpWindow
  count() -> i32
  item(index: i32) -> XHwpWindow

XHwpWindow         - 개별 HWP 창
  toolbar_layout() -> XHwpToolbarLayout
  visible() -> bool
  set_visible(bool)

XHwpToolbarLayout  - 툴바 레이아웃 관리
  create_toolbar(name, style) -> XHwpToolbar
  create_toolbar_button(name, action_id, style) -> XHwpToolbarButton
  get_toolbar(name) -> XHwpToolbar
  delete_toolbar(name) -> bool
  show_toolbar(name, show) -> bool
  create_menu_button(name, action_id, style) -> XHwpToolbarMenuButton
  get_toolbox_toolbar() -> XHncToolBoxToolbar
  change_serialize_path(path)
  is_new_serialize_path() -> bool

XHwpToolbar        - 개별 툴바
XHwpToolbarButton  - 툴바 버튼
XHwpToolbarMenuButton
XHncToolBoxToolbar - ToolBox(리본) 최상위
XHncToolBoxTab     - ToolBox 탭
XHncToolBox        - 탭 내의 ToolBox
XHncToolBoxLayout  - ToolBox 레이아웃
XHncToolBoxGroup   - ToolBox 그룹
```

---

### 액션 타입 (`h_action.rs`)

```
HAction            - 액션 실행기
  run(action_name) -> bool
  get_default(action_name, pset) -> bool
  execute(action_name, pset) -> bool
  popup_dialog(action_name, pset) -> bool

HParameterSet      - 파라미터 셋 컨테이너
  h_insert_text() -> HInsertText

HInsertText        - 텍스트 삽입 파라미터
  h_set() -> DispObj
  text() -> String
  set_text(text: &str)
```

---

### IHwpObject 열거형/구조체 (`ihwp_object.rs`)

```
enum GetTextFileFormat  { Unicode, Text, Html, Hwp, HwpMl2X }

enum ScanSpos  { Current=0x00, Specified=0x10, Line=0x20, Paragraph=0x30,
                 Section=0x40, List=0x50, Control=0x60, Document=0x70 }

enum ScanEpos  { Current=0x00, Specified=0x01, Line=0x02, Paragraph=0x03,
                 Section=0x04, List=0x05, Control=0x06, Document=0x07 }

enum ScanDirection     { Forward=0x000, Backward=0x100 }

enum GetTextStatus     { None, EndOfList, Normal, NextParagraph,
                          EnterControl, ExitControl, NotInitialized,
                          ConversionFailed, Unknown(i32) }

struct ScanRange(u32)
  new(spos: ScanSpos, epos: ScanEpos) -> Self  [const]
  backward(self) -> Self                        [const]
  within_selection() -> Self                    [const]
  as_u32(self) -> u32                           [const]

mod mask  { NORMAL=0x00, CHAR=0x01, INLINE=0x02, CTRL=0x04 }
```

---

## 액션 메서드 요약

모두 `impl HwpObject` 블록에 추가.

| 모듈 | 파일 | 메서드 수 |
|------|------|---------|
| 앱/파일/인쇄 | `actions/app.rs` | ~58 |
| 편집 | `actions/edit.rs` | ~43 |
| 찾기/바꾸기 | `actions/find.rs` | ~12 |
| 글자 모양 | `actions/chars.rs` | ~14 |
| 문단 모양 | `actions/para.rs` | ~90 |
| 커서 이동 | `actions/move_.rs` | 139 |
| 삽입 | `actions/insert.rs` | ~20 |
| 표 | `actions/table.rs` | ~211 |
| 그리기 | `actions/draw.rs` | ~168 |
| 그림 | `actions/picture.rs` | ~31 |
| 보기 | `actions/view.rs` | ~42 |
| 페이지 | `actions/page.rs` | ~16 |
| 주석 | `actions/note.rs` | ~8 |
| 변경 추적 | `actions/track.rs` | ~14 |
| 매크로 | `actions/macro_.rs` | ~51 |
| **합계** | | **600+** |

---

## 의존성

```toml
chrono      = "0.4"       # 디버그 로그 타임스탬프
thiserror   = "2.0.18"    # 에러 타입 derive
widestring  = "1.2.1"     # Wide 문자열
windows     = "0.62.2"    # Windows COM API 바인딩
  features: Win32_Foundation, Win32_System_Com, Win32_System_Ole,
            Win32_System_Variant, Win32_Globalization
```

---

## 설계 패턴

| 패턴 | 위치 | 설명 |
|------|------|------|
| Newtype macro | `hwp_types.rs` | `hwp_com_type!`으로 IDispatch 래핑 + Deref + 변환 트레이트 자동 생성 |
| Trait 기반 변환 | `variant.rs` | `FromVariant`/`IntoVariant`로 Rust 타입 ↔ VARIANT 변환 일반화 |
| Extension impl | `actions/*`, `ihwp_object.rs` | `HwpObject`에 파일별 메서드 블록 분리 추가 |
| 타입 안전 파라미터 | `ihwp_object.rs` | `ScanRange`, `ScanSpos`, `ScanEpos` 열거형으로 비트 필드 캡슐화 |

---

## 크레이트 의존 관계 (워크스페이스)

```
hwp_core  ←──  hwp_com      (OLE 클라이언트 라이브러리)
          ←──  hwp_addon     (애드인 DLL 프레임워크)
          ←──  hello_hwp_com (OLE 예제 바이너리)
          ←──  hello_hwp_dll (애드인 DLL 예제)
          ←──  hwp_dabbrev   (dabbrev 애드인)
```
