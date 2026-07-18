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

`.cargo/config.toml`에서 `i686-pc-windows-msvc` 타겟 링커를 `../../msvc-linker/linker-x86.sh`로 지정한다.  
이 스크립트는 [strickczq/msvc-wsl-rust](https://github.com/strickczq/msvc-wsl-rust) 기반이며, 내부 경로나 MSVC 버전은 환경에 맞게 수정해야 한다.

## Architecture

Two modes of interacting with HWP:

**OLE Client (`hwp_com`)** — external process connects to HWP via `CoCreateInstance` and ProgID `"HwpFrame.HwpObject.2"`.

**Add-in DLL (`hwp_addon`)** — HWP loads the DLL and calls exported `QueryUserActionInterface` to get a C-compatible vtable. Implement the `HwpUserAction` trait and use a macro to generate the export.

### Crate Graph

- **`hwp_core`** - shared foundation
  + `HwpObject` (이 프로젝트 사용자들에게 노출되는 api 진입점. IHwpObject에 대응)
  + `HwpError` (wraps `windows::core::Error`)
  + `HwpVer` (version enum: major 10→V2018, 12→V2022, 13→V2024)
- **`hwp_com`** — COM client library, depends on `hwp_core`
- **`hwp_addon`** — add-in framework, `HwpUserAction` trait, `#[repr(C)]` vtable structs
- **`hello_hwp_ole`** (pkg: `hello_hwp_com`) — example binary using OLE client approach
- **`hello_hwp_dll`** — example add-in DLL using `hwp_addon` framework
- **`hwp_dabbrev`** — hwp dynamic abbreviation addon like emacs's dabbrev.
  Sample project.

## HWP SDK Documentation

Official SDK docs are in `hwp_sdk_doc/` (2025.04 edition, gitignored):

- **`HwpAutomation_2504.pdf`** — OLE Automation API reference (methods, properties, events)
- **`ActionTable_2504.pdf`** — complete table of HWP action IDs
- **`ParameterSetTable_2504.pdf`** — parameter set definitions for actions
- **`Addon-Action.pdf`** — add-in/action plugin development guide
- **`한글오토메이션EventHandler추가_2504.pdf`** — EventHandler addition for Automation
- **`HwpCtrl.tlb.IDL`** — HwpCtrl COM 타입 라이브러리 IDL
- **`HwpObject.tlb.IDL`** — HwpObject COM 타입 라이브러리 IDL

전체 파일 목록은 `hwp_sd_doc_종류.md` 참고.

## HWP ADDON/UserAction sample in vsc++

- **`HncUserActionSample/`** - HWP ADDON(hwp_addon)에 대한 C++ sample project 
