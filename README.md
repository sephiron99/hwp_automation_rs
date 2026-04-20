# hwp_automation_rs(WIP)

한글(HWP) COM/OLE 자동화를 위한 Rust 바인딩 라이브러리.

한글의 공식 OLE Automation API를 Rust에서 안전하게 사용할 수 있도록 설계된 워크스페이스입니다. 외부 프로세스에서 한글을 제어하는 **OLE 클라이언트** 방식과, 한글 내부에서 실행되는 **애드인 DLL** 방식을 모두 지원합니다.

---

## 특징

- **두 가지 연동 방식**: OLE 클라이언트(`hwp_com`) / 애드인 DLL(`hwp_addon`)
- **600+ 액션 메서드**: 편집, 표, 그림, 찾기/바꾸기, 커서 이동, 문단 모양 등 한글 SDK 전체 액션을 Rust 메서드로 제공
- **타입 안전한 API**: `HwpError`, `HwpVer`, `ScanRange` 등 래퍼 타입으로 COM 오류를 Rust 방식으로 처리
- **버전 감지**: 실행 중인 한글 버전(V2018 / V2022 / V2024)을 자동으로 감지
- **WSL 지원**: WSL 환경에서 `i686-pc-windows-msvc` 크로스 컴파일 가능

---

## 지원 환경

| 항목 | 요구 사항 |
|------|----------|
| 운영체제 | Windows (32-bit 대상) |
| 타겟 | `i686-pc-windows-msvc` |
| 한글 버전 | V2018 (major 10), V2022 (major 12), V2024 (major 13) |
| Rust | edition 2024 |

---

## 크레이트 구조

```
hwp_automation_rs/
├── crates/
│   ├── hwp_core/        # 공유 기반 라이브러리 (HwpObject, HwpError, HwpVer, 액션 메서드)
│   ├── hwp_com/         # OLE 클라이언트 (CoCreateInstance, Running Object Table)
│   ├── hwp_addon/       # 애드인 DLL 프레임워크 (HwpUserAction 트레이트, vtable)
│   ├── hello_hwp_ole/   # OLE 클라이언트 예제
│   ├── hello_hwp_dll/   # 애드인 DLL 예제
│   └── hwp_dabbrev/     # Emacs dabbrev 스타일 자동완성 애드인
└── HncUserActionSample/ # C++ 애드인 레퍼런스 샘플
```

### 의존 관계

```
hwp_core  ←──  hwp_com      (OLE 클라이언트 라이브러리)
          ←──  hwp_addon     (애드인 DLL 프레임워크)
          ←──  hello_hwp_ole (OLE 예제 바이너리)
          ←──  hello_hwp_dll (애드인 DLL 예제)
          ←──  hwp_dabbrev   (dabbrev 애드인)
```

---

## 빌드

```bash
cargo build           # 전체 빌드
cargo build -p hwp_core  # 크레이트 개별 빌드
cargo check           # 타입 검사만
```

> 테스트는 없으며, Windows 실환경에서만 검증합니다.

### WSL 크로스 컴파일 설정

WSL에서 `i686-pc-windows-msvc` 타겟으로 빌드하려면 [strickczq/msvc-wsl-rust](https://github.com/strickczq/msvc-wsl-rust) 스크립트가 필요합니다. `.cargo/config.toml`의 `linker` 경로를 환경에 맞게 수정하세요.

---

## 사용 방법

### OLE 클라이언트 (`hwp_com`)

외부 프로세스에서 실행 중인 한글을 제어합니다.

```rust
use hwp_com::HwpClient;

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

// HwpObject 메서드 (600+ 액션)
hwp.insert_text("텍스트")?
hwp.run("FileNew")?
hwp.move_doc_begin()?
hwp.get_text_file(GetTextFileFormat::Unicode, false)?
hwp.init_scan(mask::NORMAL, ScanRange::within_selection(), 0, 0, 0, 0)?
```

---

### 애드인 DLL (`hwp_addon`)

한글이 DLL을 로드하여 내부에서 직접 실행합니다.

```rust
use hwp_addon::{export_hwp_addon, hwp_user_action::{ActionMeta, HwpUserAction, ToolbarConfig}};
use hwp_core::hwp_obj::HwpObject;

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

    fn do_action(&self, action_name: &str, hwp: &HwpObject) -> hwp_core::error::Result<bool> {
        if action_name == "MyAction" {
            hwp.insert_text("Rust 애드인에서 삽입!")?;
            return Ok(true);
        }
        Ok(false)
    }
}

export_hwp_addon!(MyPlugin, MyPlugin);
```

`Cargo.toml`에서 `crate-type = ["cdylib"]`로 설정한 뒤 빌드하면 한글에 로드할 수 있는 DLL이 생성됩니다.

---

## Hwp Addin Sample `hwp_dabbrev` — 자동완성 애드인

Emacs의 `dabbrev-expand`에서 영감을 받은 한글 자동완성 애드인입니다.

- **단축키**: `Ctrl+/`
- **prefix 기반**: 커서 앞 단어를 prefix로 삼아 문서 내 일치하는 단어를 후보로 제시
- **next-word 모드**: 커서가 단어 사이 공백에 있을 때 직전 단어 다음에 자주 나오는 단어를 제안
- **점진적 스캔**: 초기에는 인접 2개 문단만 스캔, 후보 소진 시 순차 확장
- **빈도 캐시**: 선택 이력을 기반으로 자주 쓰는 단어를 우선 제시

---

## `hwp_core` 핵심 타입

자세한 내용은 [`hwp_core.md`](hwp_core.md)를 참고하세요.

| 타입 | 설명 |
|------|------|
| `HwpObject` | 최상위 API 진입점. `IHwpObject` 대응 |
| `HwpError` | COM 오류 래퍼 (`ComError`, `ActionNotFound`, `ExecutionFailed` 등) |
| `HwpVer` | 버전 열거형 (`V2018`, `V2022`, `V2024`, `Other`) |
| `DispObj` | 일반 `IDispatch` 래퍼 (`hwp_com_type!` 매크로로 확장) |
| `ScanRange` | `InitScan` 범위 파라미터 (`ScanSpos` + `ScanEpos` 비트 필드) |
| `HAction` | 액션 실행기 (`run`, `execute`, `get_default`, `popup_dialog`) |

---

## HWP SDK 문서

공식 SDK 문서는 `hwp_sdk_doc/` 디렉토리에 있습니다 (gitignore됨). 포함 파일 목록은 [`hwp_sd_doc_종류.md`](hwp_sd_doc_종류.md)를 참고하세요.

---

## 라이선스

MIT
