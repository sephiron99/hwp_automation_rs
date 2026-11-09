# hwp_automation_rs(WIP)

한글(HWP) COM/OLE 자동화를 위한 Rust 바인딩 라이브러리.

외부 프로세스에서 한글을 제어하는 **OLE 클라이언트**(`hwp_com`)와, 한글 안에서 실행되는 **애드온 DLL**(`hwp_addon`)을 모두 지원합니다.

> **비공식 프로젝트입니다.** 주식회사 한글과컴퓨터와 무관하며, 한컴의 승인이나 지원을 받지 않았습니다.
>
> 이 라이브러리로 한글 오토메이션을 쓰는 것은 한컴의 이용 조건을 따릅니다. 한컴은 개인의 비상업적 이용만 자유롭게 허용하고, **상업적 이용에는 한컴의 별도 라이선스**를 요구합니다. 이 저장소의 MIT 라이선스는 이 저장소의 코드에만 적용됩니다. 자세한 내용은 [상표 및 권리 고지](#상표-및-권리-고지-trademark--rights-notice)를 참고하세요.

## 특징

- **편의 API와 SDK 직접 사용**: 자주 쓰는 작업은 편의 API(`hwp.ext_text()` 등)로, 그 밖의 것은 SDK method와 Action(530개 래퍼)을 그대로 호출
- **타입 안전**: 한글이 `false`로 알리는 실패까지 `HwpError`로 통일해 `?`로 처리
- **한글 2018 / 2022 / 2024**: 실행 중인 버전을 `hwp.version()`으로 확인하고, 새 버전의 API는 cargo feature로 선택
- **WSL에서 빌드**: `i686-pc-windows-msvc` 크로스 컴파일

| 항목 | 요구 사항 |
|---|---|
| 운영체제 | Windows |
| 타겟 | `i686-pc-windows-msvc`(Only) |
| 한글 버전 | 2018 (major 10), 2022 (major 12), 2024 (major 13) |
| Rust | edition 2024 |

## 빠른 시작

```toml
[dependencies]
hwp_com = "0.4"     # OLE 클라이언트
# hwp_addon = "0.4" # 애드온 DLL (crate-type = ["cdylib"])
```

```rust
use hwp_com::HwpClient;

fn main() -> hwp_com::Result<()> {
    let hwp = HwpClient::new()?;                      // 새 한글 인스턴스 시작
    hwp.windows()?.active()?.set_visible(true)?;
    hwp.ext_text().insert("Hello from Rust!")?;      // 편의 API
    hwp.h_action()?.char_shape_bold()?;              // Action
    println!("{:?}", hwp.get_pos()?);                 // SDK method
    Ok(())
}
```

애드온 DLL은 `HwpUserAction` trait을 구현하고 `export_hwp_addon!`으로 내보냅니다. 예제와 로그·버전 feature 설정은 사용법을 보세요.

## 문서

| 문서 | 내용 |
|---|---|
| [사용법](https://github.com/sephiron99/hwp_automation_rs/blob/main/사용법.md) | 크레이트 구조, OLE 클라이언트·애드온 DLL 예제, 로그, 버전 feature, 편의 API와 SDK 직접 사용, 구현 현황 |
| [Breaking changes](https://github.com/sephiron99/hwp_automation_rs/blob/main/BREAKING_CHANGES.md) | 버전을 올릴 때 고칠 코드 |
| [HWP 실전 노트](https://github.com/sephiron99/hwp_automation_rs/blob/main/hwp_실전_노트.md) | SDK 설명에 더해 알아둘 IME, Undo, 팝업, C++ ABI, 저장 동작 등의 경험 |
| [HWPML 안에서 문서 위치 찾기](https://github.com/sephiron99/hwp_automation_rs/blob/main/hwpml_위치.md) | `GetPos`의 (list, para, pos)와 HWPML 위치를 서로 바꾸는 방법 |
| [문서 전체를 텍스트로 빨리 얻는 방법](https://github.com/sephiron99/hwp_automation_rs/blob/main/문서_내보내기_속도.md) | `GetTextFile`·`SaveAs` 형식별 속도 비교 |
| [`hwp_core` 구조](https://github.com/sephiron99/hwp_automation_rs/blob/main/hwp_core.md) | 내부 층 구조와 모듈 (기여자용) |
| [미구현 SDK method](https://github.com/sephiron99/hwp_automation_rs/blob/main/bind_progress.md) · [Action 구현 계획](https://github.com/sephiron99/hwp_automation_rs/blob/main/plan.md) | 구현 현황 |
| [`hwp_dabbrev`](https://github.com/sephiron99/hwp_automation_rs/blob/main/crates/hwp_dabbrev/README.md) | Emacs dabbrev 스타일 자동완성 애드온 (예제) |

한컴의 공식 SDK 문서(Automation API, Action·ParameterSet 표, 애드온 개발 안내서, 보안 모듈 예제)는 [hancom-io/devcenter-archive](https://github.com/hancom-io/devcenter-archive)에 있습니다. method와 Action 이름은 `hwp-automation/`의 `HwpAutomation_2504.pdf`, `ActionTable_2504.pdf`, `ParameterSetTable_2504.pdf`를 기준으로 합니다.

## 빌드

```bash
cargo build              # 전체 빌드
cargo build -p hwp_core  # 크레이트 개별 빌드
cargo check              # 타입 검사만
```

테스트는 없으며, Windows 실환경에서만 검증합니다. WSL에서는 [sephiron99/msvc-wsl-rust](https://github.com/sephiron99/msvc-wsl-rust) 같은 링크 스크립트를 `link.exe`라는 이름으로 `PATH`에 두면 됩니다(예: `~/.local/bin/link.exe` 심볼릭 링크).

## 상표 및 권리 고지 (Trademark & Rights Notice)

- **비공식 프로젝트**: 이 소프트웨어는 주식회사 한글과컴퓨터(이하 한컴)와 무관한 독립적 결과물이며, 한컴의 승인·보증·지원을 받지 않았습니다.
- **참고한 자료**: 한컴이 공개한 한글 오토메이션 문서(API 설명서, Action 표, ParameterSet 표, Add-on 안내서)와 한글이 제공하는 COM 타입 라이브러리 정보를 참고해 개발했습니다. 한컴의 SDK 문서, 샘플 소스 코드, 실행 파일은 이 저장소와 크레이트에 포함하지 않습니다.
- **한글 오토메이션 이용 조건**: 한컴은 한글 오토메이션을 개인이 비상업적인 목적으로 이용하는 경우에 한해 자유롭게 쓸 수 있다고 안내하며, 상업적으로 판매되는 솔루션이나 응용 프로그램 개발 등 상업적 목적의 이용에는 한컴의 승인을 통한 별도 라이선스를 요구합니다(2026년 10월 확인). 이 라이브러리를 쓰는 프로그램에도 같은 조건이 적용됩니다. 최신 조건은 [한컴 디벨로퍼의 한글 오토메이션 안내](https://developer.hancom.com/hwpautomation)에서 직접 확인하세요.
- **상표**: "한글", "한컴", "HWP", "HWPX"는 한컴의 상표입니다. 이 프로젝트는 호환 대상을 가리키기 위해서만 이 이름을 사용합니다.

This is an unofficial project, not affiliated with or endorsed by Hancom Inc. It was developed with reference to Hancom's published HWP Automation documentation and the COM type library information exposed by HWP; no Hancom documents, sample code, or binaries are included. Hancom states that HWP Automation is free only for personal, non-commercial use and that commercial use requires a separate license from Hancom; the same applies to programs built with this library. "Hangul", "Hancom", "HWP", and "HWPX" are trademarks of Hancom Inc., used here only to identify the product this library works with.

## 라이선스

MIT. 이 저장소의 코드에만 적용되며, 한글 프로그램이나 한글 오토메이션을 이용할 권리는 포함하지 않습니다.
