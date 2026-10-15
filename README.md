# hwp_automation_rs(WIP)

한글(HWP) COM/OLE 자동화를 위한 Rust 바인딩 라이브러리.

한글의 공식 OLE Automation API를 Rust에서 안전하게 사용할 수 있도록 설계된 워크스페이스입니다. 외부 프로세스에서 한글을 제어하는 **OLE 클라이언트** 방식과, 한글 내부에서 실행되는 **애드온 DLL** 방식을 모두 지원합니다.

> **비공식 프로젝트입니다.** 주식회사 한글과컴퓨터와 무관하며, 한컴의 승인이나 지원을 받지 않았습니다.
>
> 이 라이브러리로 한글 오토메이션을 쓰는 것은 한컴의 이용 조건을 따릅니다. 한컴은 개인의 비상업적 이용만 자유롭게 허용하고, **상업적 이용에는 한컴의 별도 라이선스**를 요구합니다. 이 저장소의 MIT 라이선스는 이 저장소의 코드에만 적용됩니다. 자세한 내용은 [상표 및 권리 고지](#상표-및-권리-고지-trademark--rights-notice)를 참고하세요.

SDK 설명에 더해 알아둘 IME, Undo, 팝업, C++ ABI 등의 경험과 구현상 가정은 [HWP 실전 노트](https://github.com/sephiron99/hwp_automation_rs/blob/main/hwp_실전_노트.md)에 정리했습니다.

---

## 특징

- **두 가지 연동 방식**: OLE 클라이언트(`hwp_com`) / 애드온 DLL(`hwp_addon`)
- **편의 API와 SDK 직접 사용**: 자주 쓰는 작업은 편의 API로, 그 밖의 것은 SDK method와 Action(530개 래퍼)을 그대로 호출
- **타입 안전한 API**: `HwpVer`, `ScanRange` 등 래퍼 타입을 제공하고, 한글이 `false`로 알리는 실패까지 `HwpError`로 통일해 `?`로 처리
- **버전 확인**: `hwp.version()`으로 실행 중인 한글 버전을 `HwpVer`(V2018 / V2022 / V2024 / `Other`)로 확인. 버전을 읽지 못하면 `Other(0, 0)`
- **WSL 지원**: WSL 환경에서 `i686-pc-windows-msvc` 크로스 컴파일 가능

---

## 지원 환경

| 항목      | 요구 사항                                            |
|-----------|------------------------------------------------------|
| 운영체제  | Windows (32-bit 대상)                                |
| 타겟      | `i686-pc-windows-msvc`(Only)                         |
| 한글 버전 | V2018 (major 10), V2022 (major 12), V2024 (major 13) |
| Rust      | edition 2024                                         |

---

## 크레이트 구조

```
hwp_automation_rs/
├── crates/
│   ├── hwp_core/        # 공유 기반 라이브러리 (HwpObject, HwpError, HwpVer, SDK method·Action binding, 편의 API)
│   ├── hwp_com/         # OLE 클라이언트 (CoCreateInstance, Running Object Table)
│   ├── hwp_addon/       # 애드온 DLL 프레임워크 (HwpUserAction 트레이트, vtable)
│   ├── hello_hwp_ole/   # OLE 클라이언트 예제
│   ├── hello_hwp_dll/   # 애드온 DLL 예제
│   ├── test_hwpml_caret/ # `hwpml_with_caret` 실환경 검증용 애드온 DLL
│   └── hwp_dabbrev/     # Emacs dabbrev 스타일 자동완성 애드온
└── HncUserActionSample/ # C++ 애드온 레퍼런스 샘플
```

### 의존 관계

```
hwp_core  ←──  hwp_com    ←──  hello_hwp_ole  (OLE 예제 바이너리)
          ←──  hwp_addon  ←──  hello_hwp_dll  (애드온 DLL 예제)
                          ←──  test_hwpml_caret (hwpml_with_caret 검증용 DLL)
                          ←──  hwp_dabbrev    (dabbrev 애드온)
```

사용자는 `hwp_com`이나 `hwp_addon` 하나만 의존하면 됩니다. 두 크레이트가 `HwpObject`, `HwpError`, `Result`, `HwpExt`, `HwpVer`와 `hwp_core`의 `method`, `pset`, `actions`, `ext`, `variant`, `error`, `hwp_ver` 모듈을 같은 이름으로 다시 내보내므로, `hwp_core::method::…`는 `hwp_com::method::…`나 `hwp_addon::method::…`로 쓰면 됩니다. `hwp_core`에 직접 의존해도 같은 타입이라 섞어 써도 됩니다.

---

## 빌드

```bash
cargo build           # 전체 빌드
cargo build -p hwp_core  # 크레이트 개별 빌드
cargo check           # 타입 검사만
```

> 테스트는 없으며, Windows 실환경에서만 검증합니다.

### WSL 크로스 컴파일 설정

WSL에서 `i686-pc-windows-msvc` 타겟으로 빌드하려면
[strickczq/msvc-wsl-rust](https://github.com/strickczq/msvc-wsl-rust) 같은
wsl2-msvc 링크 스크립트가 유용합니다. `.cargo/config.toml`의 `linker` 경로를 환경에 맞게 수정하세요.

---

## 사용 방법

### OLE 클라이언트 (`hwp_com`)

외부 프로세스에서 실행 중인 한글을 제어합니다.

```rust
use hwp_com::{HwpClient, HwpExt};

fn main() -> hwp_com::Result<()> {
    // 새 한글 인스턴스 시작
    let hwp = HwpClient::new()?;
    hwp.windows()?.active()?.set_visible(true)?;

    // 실행 중인 모든 한글 인스턴스 열거
    let instances = HwpClient::list_running()?;
    for (_moniker, hwp) in &instances {
        println!("버전: {}", hwp.version().display_name());
        hwp.insert_text("Hello from Rust!")?;
    }
    Ok(())
}
```

#### 주요 API

```rust
HwpClient::new()           // 새 한글 인스턴스 시작
HwpClient::attach()        // 실행 중인 첫 번째 인스턴스에 연결
HwpClient::list_running()  // 실행 중인 모든 인스턴스 열거
```

`HwpObject`를 얻은 뒤의 사용법은 아래 [API 사용법](#api-사용법)을 참고하세요.

---

### 애드온 DLL (`hwp_addon`)

한글이 DLL을 로드하여 내부에서 직접 실행합니다.

```rust
use hwp_addon::{HwpExt, HwpObject, export_hwp_addon, hwp_user_action::{ActionMeta, HwpUserAction}};

pub struct MyPlugin;

impl HwpUserAction for MyPlugin {
    fn actions(&self) -> &'static [ActionMeta] {
        static ACTIONS: [ActionMeta; 1] = [ActionMeta {
            name: "MyAction",
            label: "내 액션",
            image_index: 0,
            shortcut: None,
        }];
        &ACTIONS
    }

    fn do_action(&self, action_name: &str, hwp: &HwpObject) -> hwp_addon::Result<bool> {
        if action_name == "MyAction" {
            hwp.insert_text("Rust 애드온에서 삽입!")?;
            return Ok(true);
        }
        Ok(false)
    }
}

export_hwp_addon!(MyPlugin, MyPlugin);
```

`Cargo.toml`에서 `crate-type = ["cdylib"]`로 설정한 뒤 빌드하면 한글에 로드할 수 있는 DLL이 생성됩니다.

### 로그

모든 크레이트는 [`tracing`](https://docs.rs/tracing) 이벤트를 내보내기만 하고, 구독자(subscriber)는 설치하지 않습니다. 아무것도 하지 않으면 로그는 어디에도 기록되지 않습니다.

- **OLE 클라이언트**: 실행 파일에서 직접 구독자를 설치합니다. 예: `tracing_subscriber::fmt().with_max_level(tracing::Level::DEBUG).init();`
- **애드온 DLL**: `hwp_addon`의 `file-log` 기능을 켜고 `on_load`에서 `hwp_addon::logging::init_default(LevelFilter::DEBUG)`를 호출하면 `%LOCALAPPDATA%\HwpAddon\hwp_addon_debug.log`에 기록합니다. 여러 번 호출해도 처음 한 번만 설치됩니다.
- **릴리스에서 아예 빼기**: 최종 바이너리/DLL 크레이트에서 `tracing`의 `release_max_level_off` 같은 기능을 켭니다. 라이브러리 크레이트에서는 켜지 마세요.

COM 호출 실패는 `debug`, 개별 COM 호출 흐름은 `trace` 레벨입니다.

### 한글 버전 선택

한글 2022, 2024에서 새로 생긴 API는 cargo feature를 켜야 쓸 수 있습니다. 아무것도 켜지 않으면 한글 2018 기준입니다.

```toml
[dependencies]
hwp_com = { version = "0.3", features = ["v2022"] } # 한글 2022 이상을 대상으로 함
```

- `v2024`는 `v2022`를 포함합니다. `hwp_addon`도 같은 이름의 feature를 씁니다.
- feature는 "이 버전 이상의 한글에서 실행한다"는 선언일 뿐이고, 실행 중인 한글 버전을 검사하지는 않습니다. 더 낮은 버전의 한글에서 그 API를 호출하면 SDK method·property는 `HwpError::UnknownMember`, Action은 `HwpError::ExecutionFailed`를 반환합니다.
- docs.rs 문서에서는 이런 API에 "Available on crate feature v2022 only" 같은 표시가 붙습니다.

아직 feature로 구분된 API는 없습니다.

---

## API 사용법

OLE 클라이언트든 애드온이든 `HwpObject`를 얻은 뒤의 사용법은 같습니다. 두 가지 방식이 있습니다.

### 1. 편의 API

자주 쓰는 작업을 한 번의 호출로 묶어 둔 것입니다. `HwpExt` trait을 가져오면 쓸 수 있습니다.

```rust
use hwp_com::HwpExt;   // 애드온은 hwp_addon::HwpExt

hwp.insert_text("텍스트")?;              // 커서 위치에 텍스트 삽입
hwp.save_as_hwpx("C:\\out.hwpx")?;       // HWPX로 저장
let shape = hwp.current_char_shape()?;  // 커서 위치의 글자 모양
let doc = hwp.hwpml_with_caret()?;      // 문서의 HWPML과 그 안의 커서 위치
let (before, after) = doc.hwpml.split_at(doc.offset);
```

`hwpml_with_caret`은 커서 자리에 책갈피를 잠깐 넣어 `GetTextFile("HWPML2X")` 출력에서 그 자리를 찾고, 실행 취소(Undo)로 책갈피를 되돌린 뒤의 HWPML을 돌려줍니다. `GetPos` 값과 HWPML의 `CARETPOS`는 리스트·문단·글자 번호일 뿐이고, HWPML 본문 요소에는 그 번호에 대응하는 속성이 없어 커서가 HWPML의 어디인지 알 수 없기 때문입니다. 문서를 잠깐 수정하므로 수정 상태와 다시 실행(Redo) 기록이 바뀌고, `GetTextFile`을 두 번 부르므로 큰 문서에서는 느립니다. 선택 블록이 있으면 선택 방향과 상관없이 블록의 앞쪽 끝을 돌려주고, 호출 뒤에는 선택이 풀리며 커서가 블록의 뒤쪽 끝으로 갑니다.

애드온에서는 `hwp_addon`의 `HwpTextExt`(문서 텍스트 순회), `HwpEditExt`(커서 앞 문자 교체)도 같은 방식으로 씁니다.

### 2. SDK 직접 사용

편의 API에 없는 기능은 한글 SDK의 method와 Action을 그대로 호출합니다. 한컴 SDK 문서의 이름과 1:1로 대응합니다.

**method** — `HwpAutomation` 문서의 method/property. `HwpObject`와 하위 객체에 직접 붙어 있습니다.

```rust
use hwp_com::method::hwp_object::{GetTextFileFormat, ScanRange, mask};   // 애드온은 hwp_addon::method::…

let (list, para, pos) = hwp.get_pos()?;
hwp.get_text_file(GetTextFileFormat::Unicode, false)?;
hwp.init_scan(mask::NORMAL, ScanRange::within_selection(), 0, 0, 0, 0)?;
hwp.windows()?.active()?.set_visible(true)?;
```

**Action** — `ActionTable` 문서의 Action ID. `hwp.act()` 뒤에 있습니다.

```rust
hwp.act().undo()?;              // Action ID "Undo"
hwp.act().char_shape_bold()?;   // Action ID "CharShapeBold"
```

래퍼가 아직 없는 것은 이름으로 직접 호출할 수 있습니다.

```rust
hwp.act().run("FileNew")?;                            // 래퍼가 없는 Action
let n: i32 = hwp.get("PageCount")?;                   // 래퍼가 없는 property
let ok: bool = hwp.call_with("Open", vec![/* .. */])?; // 래퍼가 없는 method
```

파라미터가 필요한 Action은 `hwp.h_action()`과 `hwp.h_parameter_set()`으로 실행합니다. Action을 인자 없이 실행하면 한글이 대화상자를 띄울 수 있습니다.

---

## 구현 현황

아직 SDK 전체를 덮지 못했습니다. 래퍼가 없는 기능은 위의 이름 기반 호출로 쓸 수 있습니다.

| 대상 | 현황 | 자세한 내용 |
|---|---|---|
| SDK method | IDL의 366개 중 31개 | [`bind_progress.md`](https://github.com/sephiron99/hwp_automation_rs/blob/main/bind_progress.md) — 미구현 목록 |
| Action | 530개 래퍼 (파라미터 없는 실행만) | [`plan.md`](https://github.com/sephiron99/hwp_automation_rs/blob/main/plan.md) — 전체 Action 구현 계획 |
| 이벤트 (`IHwpObjectEvents`) | 미구현 | [`bind_progress.md`](https://github.com/sephiron99/hwp_automation_rs/blob/main/bind_progress.md) |

---

## Hwp Addon Sample `hwp_dabbrev` — 자동완성 애드온

Emacs의 `dabbrev-expand`에서 영감을 받은 한글 자동완성 애드온입니다.

- **단축키**: `Ctrl+/`
- **prefix 기반**: 커서 앞 단어를 prefix로 삼아 문서 내 일치하는 단어를 후보로 제시
- **next-word 모드**: 커서가 단어 사이 공백에 있을 때 문서에서 직전 단어 다음에 등장하는 단어를 제안
- **매 호출 전체 스캔**: 액션마다 문서 전체에서 단어 목록을 새로 구성하고, 중복을 제거한 모든 후보를 팝업에 전달
- **선택 이력 우선**: 팝업에서 확정한 단어는 다음 호출부터 후보 앞쪽에 표시(확정 횟수가 많은 순, 같으면 최근 순). 나머지 후보는 문서 등장 순서를 유지. 이력은 한글이 실행 중인 동안만 유지하며 모든 문서가 공유
- **후보 선택**: 팝업에서 `↑` / `↓` / `Ctrl+/`로 순환, `Enter`로 확정, `Esc`로 취소
- **후보 현황**: 팝업 하단에 현재 선택 번호와 전체 후보 수를 표시하고, 화면 밖에 후보가 더 있으면 해당 방향의 `↑` / `↓` 표시

---

## `hwp_core` 핵심 타입

`HwpObject`, `HwpError`, `HwpVer`, `HwpExt`는 `hwp_com`, `hwp_addon` 루트에서도 같은 이름으로 쓸 수 있습니다. 자세한 내용은 [`hwp_core.md`](https://github.com/sephiron99/hwp_automation_rs/blob/main/hwp_core.md)를 참고하세요.

| 타입 | 설명 |
|------|------|
| `HwpObject` | 최상위 API 진입점. `IHwpObject` 대응 |
| `HwpError` | COM 오류 래퍼 (`ComError`, `ActionNotFound`, `ExecutionFailed` 등) |
| `HwpVer` | 버전 열거형 (`V2018`, `V2022`, `V2024`, `Other`) |
| `DispObj` | 일반 `IDispatch` 래퍼 (`hwp_com_type!` 매크로로 확장) |
| `ScanRange` | `InitScan` 범위 파라미터 (`ScanSpos` + `ScanEpos` 비트 필드) |
| `HAction` | 액션 실행기 (`run`, `execute`, `get_default`, `popup_dialog`) |
| `Act` | Action 래퍼 핸들 (`hwp.act()`) |
| `HwpExt` | 편의 API trait |

---

## Breaking changes

### 0.3.0

**로그가 `tracing`으로 바뀌었습니다.** `hwp_core::debug` 모듈과 `hwp_addon::debug::log`가 없어졌습니다. 로그를 남기려면 `tracing::debug!` 등을 쓰고, 파일로 받으려면 위 [로그](#로그) 절을 참고하세요. 더 이상 라이브러리가 `%LOCALAPPDATA%`에 로그 파일을 저절로 만들지 않습니다. `hwp_addon::debug::{msgbox, msgbox_err}`는 그대로입니다.

### 0.2.0

0.1.0에서 올리면 아래 네 가지를 고쳐야 합니다.

**1. Action 호출은 `hwp.act()`를 거칩니다.**

```rust
// 0.1.0
hwp.undo()?;
hwp.char_shape_bold()?;

// 0.2.0
hwp.act().undo()?;
hwp.act().char_shape_bold()?;
```

`select()`, `move_word_begin()`을 포함해 Action ID에 대응하는 메서드 전부가 해당됩니다. `hwp.get_pos()` 같은 SDK method는 그대로입니다.

**2. 편의 API는 `HwpExt` trait을 가져와야 쓸 수 있습니다.**

`insert_text`, `save_as_hwpx`, `current_char_shape`, `current_para_shape`, `hwpml_with_caret`이 해당됩니다.

```rust
use hwp_core::ext::HwpExt;   // 또는 hwp_com::HwpExt
```

**3. `hwp_core`의 모듈 경로가 바뀌었습니다.**

| 0.1.0 | 0.2.0 |
|---|---|
| `hwp_core::ihwpobject::lib` | `hwp_core::method::hwp_object` |
| `hwp_core::ihwpobject::movepos` | `hwp_core::method::movepos` |
| `hwp_core::ihwpobject::{CharShape, ParaShape}` | `hwp_core::ext::{CharShape, ParaShape}` |
| `hwp_core::hwp_types` | `hwp_core::method::xhwp` |
| `hwp_core::h_action::HAction` | `hwp_core::method::h_action::HAction` |
| `hwp_core::h_action::{HParameterSet, HCharShape, HParaShape, HInsertText}` | `hwp_core::pset::{..}` |

**4. 성공 여부를 `bool`로 돌려주던 메서드가 `Result<()>`를 돌려줍니다.**

한글이 `false`(실패)를 돌려주면 `Ok(false)` 대신 `Err(HwpError::ExecutionFailed)`가 됩니다. 반환값을 검사하던 코드는 `?`로 바꾸면 됩니다.

| 대상 | 메서드 |
|---|---|
| `HwpObject` | `save_as`, `init_scan`, `select_text` |
| `HwpExt` | `save_as_hwpx` |
| `XHwpToolbarLayout` | `delete_toolbar`, `show_toolbar` |
| `XHwpToolbar` | `delete_button` |

```rust
// 0.1.0
if !hwp.save_as(path, "HWPX", "")? { /* 실패 처리 */ }

// 0.2.0
hwp.save_as(path, "HWPX", "")?;
```

`key_indicator`의 실패도 `ComError`가 아니라 `ExecutionFailed`로 바뀌었습니다. `is_new_serialize_path`, `visible`처럼 `bool`이 질문의 답인 메서드는 그대로 `Result<bool>`입니다.

---

## HWP SDK 문서

공식 SDK 문서는 `hwp_sdk_doc/` 디렉토리에 있습니다 (gitignore됨). 포함 파일 목록은 [`hwp_sd_doc_종류.md`](https://github.com/sephiron99/hwp_automation_rs/blob/main/hwp_sd_doc_종류.md)를 참고하세요.

---

## 상표 및 권리 고지 (Trademark & Rights Notice)

- **비공식 프로젝트**: 이 소프트웨어는 주식회사 한글과컴퓨터(이하 한컴)와 무관한 독립적 결과물이며, 한컴의 승인·보증·지원을 받지 않았습니다.
- **참고한 자료**: 한컴이 공개한 한글 오토메이션 문서(API 설명서, Action 표, ParameterSet 표, Add-on 안내서)와 한글이 제공하는 COM 타입 라이브러리 정보를 참고해 개발했습니다. 한컴의 SDK 문서, 샘플 소스 코드, 실행 파일은 이 저장소와 크레이트에 포함하지 않습니다.
- **한글 오토메이션 이용 조건**: 한컴은 한글 오토메이션을 개인이 비상업적인 목적으로 이용하는 경우에 한해 자유롭게 쓸 수 있다고 안내하며, 상업적으로 판매되는 솔루션이나 응용 프로그램 개발 등 상업적 목적의 이용에는 한컴의 승인을 통한 별도 라이선스를 요구합니다(2026년 10월 확인). 이 라이브러리를 쓰는 프로그램에도 같은 조건이 적용됩니다. 최신 조건은 [한컴 디벨로퍼의 한글 오토메이션 안내](https://developer.hancom.com/hwpautomation)에서 직접 확인하세요.
- **상표**: "한글", "한컴", "HWP", "HWPX"는 한컴의 상표입니다. 이 프로젝트는 호환 대상을 가리키기 위해서만 이 이름을 사용합니다.

This is an unofficial project, not affiliated with or endorsed by Hancom Inc. It was developed with reference to Hancom's published HWP Automation documentation and the COM type library information exposed by HWP; no Hancom documents, sample code, or binaries are included. Hancom states that HWP Automation is free only for personal, non-commercial use and that commercial use requires a separate license from Hancom; the same applies to programs built with this library. "Hangul", "Hancom", "HWP", and "HWPX" are trademarks of Hancom Inc., used here only to identify the product this library works with.

## 라이선스

MIT. 이 저장소의 코드에만 적용되며, 한글 프로그램이나 한글 오토메이션을 이용할 권리는 포함하지 않습니다.
