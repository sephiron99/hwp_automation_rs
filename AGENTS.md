## Project Overview

Rust workspace for controlling 한글(HWP, Hangul Word Processor) via Windows COM/OLE automation. Targets 32-bit Windows (`i686-pc-windows-msvc`).
프로젝트의 최종적인 목표는 한글 COM/OLE automation에 대해 완벽한 Rust bindings을
만드는 것이다. 
한글로 대답해.

## Build Commands

```bash
cargo build
cargo build -p hwp_core # single crate
cargo check             # type-check only
```
테스트 없음, Windows 실환경에서만 검증

## WSL 크로스 컴파일 링커 설정

저장소의 `.cargo/config.toml`은 링커를 지정하지 않으므로, rustc 기본값대로 `PATH`의 `link.exe`로 링크한다.  
WSL에서는 [sephiron99/msvc-wsl-rust](https://github.com/sephiron99/msvc-wsl-rust) 기반 링커 스크립트를 `link.exe`라는 이름으로 `PATH`에 둔다(예: `~/.local/bin/link.exe` 심볼릭 링크). 스크립트 내부 경로나 MSVC 버전은 환경에 맞게 수정해야 한다. `cargo check`는 링커 없이 동작한다.

## Architecture

Two modes of interacting with HWP:

**OLE Client (`hwp_com`)** — external process connects to HWP via `CoCreateInstance` and ProgID `"HwpFrame.HwpObject.2"`.

**Add-on DLL (`hwp_addon`)** — HWP loads the DLL and calls exported `QueryUserActionInterface` to get a C-compatible vtable. Implement the `HwpUserAction` trait and use a macro to generate the export.

### Crate Graph

- **`hwp_core`** - shared foundation
  + `HwpObject` (이 프로젝트 사용자들에게 노출되는 api 진입점. IHwpObject에 대응)
  + `HwpError` (wraps `windows::core::Error`)
  + `HwpVer` (version enum: major 10→V2018, 12→V2022, 13→V2024)
  + 내부는 4층이고 아래에서 위로만 의존한다: `pset` ← `method` ← `actions` ← `ext` (아래 "hwp_core 층 규칙" 참고)
- **`hwp_com`** — COM client library, depends on `hwp_core`
- **`hwp_addon`** — add-on framework, `HwpUserAction` trait, `#[repr(C)]` vtable structs
- **`hello_hwp_ole`** (pkg: `hello_hwp_com`) — example binary using OLE client approach
- **`hello_hwp_dll`** — example add-on DLL using `hwp_addon` framework
- **`test_hwpml_caret`** — add-on DLL for verifying `ExtDoc::hwpml_with_caret` in a real HWP
- **`hwp_dabbrev`** — hwp dynamic abbreviation addon like emacs's dabbrev.
  Sample project.

### hwp_core 층 규칙

| 층 | 모듈 | 내용 | 사용자 호출 형태 |
|---|---|---|---|
| 1 | `pset` | ParameterSet 타입 (`HParameterSet`, `HCharShape` 등) | 타입으로만 등장 |
| 2 | `method` | `HwpObject.tlb.IDL`의 method/property 1:1 binding | `hwp.get_pos()` |
| 3 | `actions` | ActionTable의 Action ID별 래퍼 (`HAction`의 메서드) | `hwp.h_action()?.undo()` |
| 4 | `ext` | 편의 API (`ext_` 접두어 핸들) | `hwp.ext_text().insert(..)` |

- 1~3층은 호출 하나에 COM 호출 하나를 대응시키고, 실패는 공통 에러 `HwpError`로 돌려준다. SDK가 성공 여부를 `BOOL`로 돌려주면 `false`를 `HwpError::ExecutionFailed`로 바꾼다. `BOOL`이 질문의 답이면 `Result<bool>`로 둔다. 에러를 삼키거나 기본값으로 대체하지 않는다(그런 처리는 `ext`). 인자를 enum으로 감싸는 것은 허용.
- 호출을 둘 이상 묶거나, 기본값을 정하거나, 에러를 삼키는 코드는 `ext`에 둔다. `ext`는 아래 층의 공개 API만 쓴다.
- 구조를 건드린 뒤에는 `cargo check`와 `scripts/check_layers.sh`(역방향 참조 검사)를 실행한다.
- 한글 2022·2024에서 새로 생긴 API는 `#[cfg(feature = "v2022")]`/`v2024`로 감싼다. feature로 감싼 코드를 바꾼 뒤에는 `cargo check --all-features`도 실행한다. 자세한 규칙은 `hwp_core.md`의 "버전 feature".
- 사용자 문서(`README.md`)에서는 4층을 드러내지 않고 "편의 API"와 "SDK 직접 사용(method, Action)" 두 갈래로만 설명한다.
- 자세한 구조는 `hwp_core.md`, SDK method 미구현 목록은 `bind_progress.md`, Action 구현 계획은 `plan.md`.

## HWP SDK Documentation

한컴이 공개한 SDK 문서(2025.04판)의 원본은 [hancom-io/devcenter-archive](https://github.com/hancom-io/devcenter-archive)에 있다.
같은 파일을 `hwp_sdk_doc/`(gitignored)에 받아 두고 읽는다. IDL 두 개와 HWPML 명세는 archive에 없다.

- **`HwpAutomation_2504.pdf`** — OLE Automation API reference (methods, properties, events)
- **`ActionTable_2504.pdf`** — complete table of HWP action IDs
- **`ParameterSetTable_2504.pdf`** — parameter set definitions for actions
- **`Addon-Action.pdf`** — add-on/action plugin development guide
- **`한글오토메이션EventHandler추가_2504.pdf`** — EventHandler addition for Automation
- **`HwpCtrl.tlb.IDL`** — HwpCtrl COM 타입 라이브러리 IDL
- **`HwpObject.tlb.IDL`** — HwpObject COM 타입 라이브러리 IDL
- **`한글문서파일형식3.0_HWPML_revision1.2.pdf`** — HWPML 파일 형식 명세
- **`보안모듈(Automation)/`** — 파일 경로 검사 보안 모듈(`RegisterModule("FilePathCheckDLL", ..)`) 예제: C++ 소스, 빌드된 32비트 DLL, 레지스트리 등록 화면

archive 경로와 파일별 출처는 `hwp_sd_doc_종류.md` 참고.

## HWP ADDON/UserAction sample in vsc++

- **`HncUserActionSample/`** - HWP ADDON(hwp_addon)에 대한 C++ sample project 
