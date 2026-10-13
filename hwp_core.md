# hwp_core 구조도

## 크레이트 개요

Windows COM/OLE를 통해 한글(HWP)을 Rust에서 제어하는 공유 라이브러리. 모든 COM 호출, 타입 변환, SDK method·Action binding을 캡슐화하며 `hwp_com` 및 `hwp_addon`은 hwp_core 크레이트를 기반으로 한다.

---

## 층 구조

`hwp_core`는 네 층으로 나뉘고, 의존은 아래에서 위로만 향한다.

```
기반  ←  pset  ←  method  ←  actions  ←  ext
```

| 층 | 모듈 | 내용 | 쓸 수 있는 것 |
|---|---|---|---|
| 기반 | `error`, `hwp_ver`, `variant`, `com_util`, `disp_obj`, `hwp_obj` | `HwpObject` 구조체, 범용 `get`/`call`/`call_with`, `hwp_com_type!` | — |
| 1 | `pset` | ParameterSet 타입 | 기반 |
| 2 | `method` | `HwpObject.tlb.IDL` interface의 method/property 1:1 binding | 기반, `pset` |
| 3 | `actions` | ActionTable의 Action ID별 래퍼 (`Act` 핸들) | 기반, `pset`, `method` |
| 4 | `ext` | 편의 API (`HwpExt` trait) | 아래 층의 공개 API만 |

### 어느 층에 둘 것인가

- **1~3층은 호출 하나에 COM 호출 하나를 대응시키고, 실패는 공통 에러 `HwpError`로 돌려준다.**
  SDK가 성공 여부를 `BOOL`로 돌려주면 `false`를 `HwpError::ExecutionFailed`로 바꾼다(`error::ensure` 사용).
  `BOOL`이 질문의 답이면(`is_new_serialize_path`, `visible` 등) `Result<bool>`로 둔다.
  에러를 삼키거나 기본값으로 대체하지 않는다.
  인자를 enum으로 감싸는 것(`MovePos`, `ScanRange`)은 타입만 입힌 것이므로 허용한다.
- **호출을 둘 이상 묶거나, 기본값을 정하거나, 에러를 삼키면 4층(`ext`)이다.**
- ParameterSet은 method와 Action이 함께 쓰므로(`GetPosBySet`, `ExportStyle` 등) 어느 한쪽이 아니라 `pset`에 둔다.
- Action 래퍼는 `HAction.Run`(SDK method)을 호출하므로 `actions`가 `method` 위에 있다.

`scripts/check_layers.sh`가 역방향 참조(`pset`에서 `crate::method` 등)를 찾아 실패시킨다.
구조를 바꾼 뒤에는 `cargo check`와 함께 실행한다.

사용자 문서(`README.md`)에서는 이 네 층을 드러내지 않고 "편의 API"와 "SDK 직접 사용" 두 갈래로만 설명한다.

### re-export 정책

사용자가 `hwp_core`에 따로 의존하지 않도록 `hwp_com`과 `hwp_addon`은 루트에서 같은 목록을 re-export한다.

- 항목: `HwpObject`, `error::{HwpError, Result}`, `ext::HwpExt`, `hwp_ver::HwpVer`
- 모듈: `error`, `hwp_ver`, `variant`, `pset`, `method`, `actions`, `ext`
- 제외하는 것: `com_util`, `disp_obj`, `hwp_obj` (내부용. `HwpObject`는 루트 re-export로 충분)
- `#[doc(no_inline)]`을 붙여 문서가 복사되지 않고 `hwp_core` 링크로만 나오게 한다.
- `hwp_core`에 공개 모듈을 새로 만들면 두 목록에 넣을지 정한다. `hwp_addon`에는 위 모듈과 같은 이름의 모듈을 만들지 않는다.

---

## 모듈 트리

```
hwp_core (lib.rs)
│
│  ── 기반 ──
├── error          - HwpError, Result<T>
├── hwp_ver        - HwpVer (버전 열거형)
├── variant        - FromVariant / IntoVariant 트레이트
├── com_util       - 저수준 COM Invoke 래퍼
├── disp_obj       - DispObj (일반 IDispatch 래퍼), hwp_com_type! 매크로
├── hwp_obj        - HwpObject (최상위 API 진입점)
│
│  ── 1. ParameterSet 타입 ──
├── pset           - HParameterSet, HCharShape, HParaShape, HInsertText
│
│  ── 2. SDK method ──
├── method
│   ├── hwp_object - IHwpObject 메서드, 열거형, ScanRange
│   ├── movepos    - IHwpObject::MovePos
│   ├── h_action   - HAction
│   └── xhwp       - IXHwp* 하위 객체 (문서·창·툴바)
│
│  ── 3. Action ──
├── actions        - Act 핸들 (hwp.act()) + 도메인별 래퍼
│   ├── app        - 앱/파일/인쇄/프레젠테이션/버전
│   ├── edit       - 편집/클립보드/선택/실행취소
│   ├── find       - 찾기/바꾸기
│   ├── chars      - 글자 모양
│   ├── para       - 문단 모양/스타일/글머리표
│   ├── move_      - 커서 이동
│   ├── insert     - 삽입
│   ├── table      - 표 조작
│   ├── draw       - 그리기 개체/도형
│   ├── picture    - 그림 속성/효과
│   ├── view       - 보기 옵션/화면 배율
│   ├── page       - 편집 용지/바탕쪽/구역
│   ├── note       - 각주/미주/주석
│   ├── track      - 변경 추적
│   └── macro_     - 매크로/빠른 교정/빠른 책갈피
│
│  ── 4. 편의 API ──
└── ext            - HwpExt trait, CharShape, ParaShape
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

+ method/hwp_object.rs, method/movepos.rs impl (SDK method):
  run(action: &str) -> Result<()>
  h_action() -> Result<HAction>
  h_parameter_set() -> Result<HParameterSet>
  documents() -> Result<XHwpDocuments>
  windows() -> Result<XHwpWindows>
  path() -> Result<String>
  get_text_file(format, save_block) -> Result<String>
  set_text_file(data, format, option) -> Result<i32>
  save_as(path, format, arg) -> Result<()>
  init_scan(option, range, spara, spos, epara, epos) -> Result<()>
  get_text() -> Result<(GetTextStatus, String)>
  release_scan() -> Result<()>
  get_pos() -> Result<(i32, i32, i32)>
  select_text(spara, spos, epara, epos) -> Result<()>
  key_indicator() -> Result<KeyIndicatorInfo>
  move_pos(to: MovePos) -> Result<()>

+ actions/mod.rs impl:
  act() -> Act<'_>            (Action 래퍼 핸들: hwp.act().undo()?, hwp.act().run("ActionId")?)

+ ext::HwpExt trait (use hwp_core::ext::HwpExt 또는 hwp_com::HwpExt / hwp_addon::HwpExt 필요):
  insert_text(text: &str) -> Result<()>
  save_as_hwpx(path: &str) -> Result<()>
  current_char_shape() -> Result<CharShape>
  current_para_shape() -> Result<ParaShape>
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

### COM 타입 구조체 (`method/xhwp.rs`)

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
  delete_toolbar(name) -> ()
  show_toolbar(name, show) -> ()
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

### 액션 실행기 (`method/h_action.rs`)

```
HAction            - 액션 실행기
  run(action_name) -> Result<()>
  get_default(action_name, pset) -> Result<()>
  execute(action_name, pset) -> Result<()>
  popup_dialog(action_name, pset) -> Result<()>
```

### ParameterSet 타입 (`pset/mod.rs`)

```
HParameterSet      - 파라미터 셋 컨테이너
  h_insert_text() -> HInsertText
  h_char_shape() -> HCharShape
  h_para_shape() -> HParaShape

HInsertText        - 텍스트 삽입 파라미터
  h_set() -> DispObj
  text() -> String
  set_text(text: &str)

HCharShape, HParaShape
  h_set() -> DispObj
```

---

### IHwpObject 열거형/구조체 (`method/hwp_object.rs`)

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

## Action 래퍼 요약

모두 `impl Act<'_>` 블록에 있고, `hwp.act()`로 얻은 핸들에서 호출한다.
본문은 전부 `self.0.h_action()?.run("ActionId")` 한 형태다.
전용 래퍼가 없는 Action은 `hwp.act().run("ActionId")`로 실행한다.

| 모듈 | 파일 | 메서드 수 |
|------|------|---------|
| 앱/파일/인쇄 | `actions/app.rs` | 37 |
| 편집 | `actions/edit.rs` | 34 |
| 찾기/바꾸기 | `actions/find.rs` | 12 |
| 글자 모양 | `actions/chars.rs` | 13 |
| 문단 모양 | `actions/para.rs` | 53 |
| 커서 이동 | `actions/move_.rs` | 1 |
| 삽입 | `actions/insert.rs` | 19 |
| 표 | `actions/table.rs` | 114 |
| 그리기 | `actions/draw.rs` | 107 |
| 그림 | `actions/picture.rs` | 24 |
| 보기 | `actions/view.rs` | 31 |
| 페이지 | `actions/page.rs` | 22 |
| 주석 | `actions/note.rs` | 8 |
| 변경 추적 | `actions/track.rs` | 15 |
| 매크로 | `actions/macro_.rs` | 40 |
| **합계** | | **530** |

ActionTable 전체를 덮는 계획은 `plan.md`, SDK method 쪽 미구현 목록은 `bind_progress.md` 참고.

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
| Newtype macro | `disp_obj.rs` | `hwp_com_type!`으로 IDispatch 래핑 + Deref + 변환 트레이트 자동 생성 |
| Trait 기반 변환 | `variant.rs` | `FromVariant`/`IntoVariant`로 Rust 타입 ↔ VARIANT 변환 일반화 |
| Extension impl | `method/*` | `HwpObject`에 파일별 메서드 블록 분리 추가 |
| 핸들 타입 | `actions/*` | `Act<'_>`에 Action 래퍼를 모아 `HwpObject` 네임스페이스와 분리 |
| Extension trait | `ext` | `HwpExt`로 편의 API를 SDK binding과 분리 |
| 타입 안전 파라미터 | `method/hwp_object.rs`, `method/movepos.rs` | `ScanRange`, `ScanSpos`, `ScanEpos` 열거형으로 비트 필드 캡슐화 |

---

## 크레이트 의존 관계 (워크스페이스)

```
hwp_core  ←──  hwp_com      (OLE 클라이언트 라이브러리)
          ←──  hwp_addon     (애드온 DLL 프레임워크)
          ←──  hello_hwp_com (OLE 예제 바이너리)
          ←──  hello_hwp_dll (애드온 DLL 예제)
          ←──  hwp_dabbrev   (dabbrev 애드온)
```
