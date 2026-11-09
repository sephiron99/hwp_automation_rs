# Breaking changes

버전을 올릴 때 고쳐야 할 코드를 버전별로 정리했습니다. 위가 최신입니다.

## 0.4.0

**Action 래퍼가 `HAction`으로 옮겨졌습니다.** `HwpObject::act()`와 `Act` 타입이 없어졌습니다.

```rust
// 0.3.0
hwp.act().undo()?;
hwp.act().run("FileNew")?;

// 0.4.0
hwp.h_action()?.undo()?;
hwp.h_action()?.run("FileNew")?;
```

**`HwpError::OleInitFailed`가 없어졌습니다.** OLE 초기화 실패는 원래 HRESULT가 담긴 `HwpError::ComError`로 옵니다(예: 이미 MTA로 초기화된 스레드의 `RPC_E_CHANGED_MODE`). 이 variant를 `match`하던 코드는 지우면 됩니다.

`HwpClient` 사용법은 그대로입니다. 이전에는 `new`, `attach`, `list_running`이 부를 때마다 `OleInitialize`를 부르고 해제하지 않았는데, 이제 스레드마다 처음 한 번만 초기화하고 스레드가 끝날 때 해제합니다. 해제 시점을 직접 정하려면 새로 생긴 `hwp_com::OleGuard`를 쓰면 됩니다.

여러 Action을 이어서 실행할 때는 `let action = hwp.h_action()?;`로 한 번 얻어 두고 `action.undo()?` 처럼 부르면 Action마다 `HAction`을 다시 얻지 않습니다.

**편의 API가 `ext_` 핸들로 옮겨졌습니다.** `HwpExt` trait과 `hwp_com::HwpExt`, `hwp_addon::HwpExt` 재내보내기가 없어졌고, `hwp_addon::text_extract`(`HwpTextExt`, `TextWalker`, `TextSegment`)는 `hwp_core::ext`로 옮겨져 OLE 클라이언트에서도 쓸 수 있습니다.

| 0.3.0 | 0.4.0 |
|---|---|
| `hwp.insert_text(t)` | `hwp.ext_text().insert(t)` |
| `hwp.text_segments()` / `text_segments_with(..)` | `hwp.ext_text().segments()` / `segments_with(..)` |
| `hwp.get_all_text()` | `hwp.ext_text().all()` |
| `hwp.visit_all_text(f)` | `hwp.ext_text().visit(f)` |
| `hwp.current_char_shape()` / `current_para_shape()` | `hwp.ext_shape().current_char()` / `current_para()` |
| `hwp.save_as_hwpx(p)` | `hwp.ext_doc().save_as_hwpx(p)` |
| `hwp.hwpml_with_caret()` | `hwp.ext_doc().hwpml_with_caret()` |
| `hwp_addon::text_extract::TextSegment` | `hwp_core::ext::TextSegment` (`hwp_addon::ext::TextSegment`) |

`use …::HwpExt;`와 `use hwp_addon::text_extract::HwpTextExt;`는 지우면 됩니다.

**`save_as`의 형식 인자가 `FileFormat`입니다.** 형식 이름을 문자열 대신 `FileFormat` 상수로 넘깁니다. 상수에 없는 형식은 `FileFormat::new("이름")`으로 만듭니다.

```rust
// 0.3.0
hwp.save_as(path, "HWPX", "")?;

// 0.4.0
use hwp_com::method::hwp_object::FileFormat;
hwp.save_as(path, &FileFormat::HWPX, "")?;
```

**`XHwpDocument::edit_mode()`와 `modified()`의 반환 타입이 바뀌었습니다.** 한글이 돌려주는 정수 값을 그대로 감싼 타입입니다.

| 0.3.0 | 0.4.0 |
|---|---|
| `edit_mode() -> Result<i32>` | `edit_mode() -> Result<EditMode>`. 배포용 문서는 다른 모드와 합쳐져 오므로(실측 17) `contains(EditMode::DISTRIBUTION)`으로, 읽기 전용은 `is_read_only()`로 판단합니다. 정수는 `bits()` |
| `modified() -> Result<bool>` | `modified() -> Result<ModifiedState>`. 변경 여부는 `is_modified()`. 0.3.0의 `bool` 버전은 한글이 정수를 돌려줘 항상 `VariantConversion` 에러가 났습니다 |

**문서 안의 위치가 타입으로 바뀌었습니다.** 새 모듈 `pos`(`hwp_com::pos`, `hwp_addon::pos`)의 `ListParaPos { list, para, pos }`, 같은 리스트 안의 `ParaPos { para, pos }`, 범위 `ListParaRange { start, end }`를 씁니다. `SelectText`·`InitScan`·`MovePos`처럼 리스트 없이 받는 SDK 함수는 `ParaPos`를 받습니다.

| 0.3.0 | 0.4.0 |
|---|---|
| `let (list, para, pos) = hwp.get_pos()?;` | `let p = hwp.get_pos()?;` → `p.list`, `p.para`, `p.pos` |
| `hwp.select_text(spara, spos, epara, epos)?` | `hwp.select_text(ParaPos::new(spara, spos), ParaPos::new(epara, epos))?` (현재 위치에서 시작하면 `hwp.get_pos()?.para_pos()`) |
| `hwp.init_scan(option, range, spara, spos, epara, epos)?` | `hwp.init_scan(option, range, ParaPos::new(spara, spos), ParaPos::new(epara, epos))?` |
| `hwp.text_segments_with(option, range, 0, 0, 0, 0)` | `hwp.ext_text().segments_with(option, range, ParaPos::new(0, 0), ParaPos::new(0, 0))` |
| `MovePos::Main { para, pos }` / `MovePos::CurList { para, pos }` | `MovePos::Main(ParaPos::new(para, pos))` / `MovePos::CurList(..)` |

## 0.3.0

**로그가 `tracing`으로 바뀌었습니다.** `hwp_core::debug` 모듈과 `hwp_addon::debug::log`가 없어졌습니다. 로그를 남기려면 `tracing::debug!` 등을 쓰고, 파일로 받으려면 [사용법의 로그](사용법.md#로그) 절을 참고하세요. 더 이상 라이브러리가 `%LOCALAPPDATA%`에 로그 파일을 저절로 만들지 않습니다. `hwp_addon::debug::{msgbox, msgbox_err}`는 그대로입니다.

## 0.2.0

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
