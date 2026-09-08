# HWP 자동화 개발 중 얻은 경험 지식

SDK의 API 표만으로 알기 어려웠던 동작과 이 프로젝트에서 사용한 우회 방법을 정리한다. 같은 문제를 다시 조사하지 않도록 증상, 대응 방법, 근거, 확인 범위를 함께 남긴다.

정리 기준은 **2026-09-07, 커밋 `53d0908`**이다. 코드 주석과 Git 이력을 조사하고, 로컬 SDK의 `HwpAutomation_2504.pdf`와 `Addon-Action.pdf`에서 관련 설명을 대조했다. 이번 문서 작성 중 Windows에서 새로 재현한 것은 아니다.

다음 세 종류를 구분한다.

| 구분 | 의미 |
|---|---|
| 관찰 기록 | 기존 코드 주석이나 수정 이력에 실제 증상·진단 결과가 남아 있음 |
| 구현 선택 | 해당 문제를 다루기 위해 현재 코드가 채택한 방법. HWP의 공식 계약은 아님 |
| 미확인 가정 | 구현이 의존하지만 버전별 재현이나 명확한 계약이 확인되지 않음 |

IME 진단 주석에는 HWP 2022/2024가 명시돼 있다. 다른 항목은 정확한 HWP 빌드, Windows 빌드, 입력기 버전이 기록돼 있지 않으므로 모든 환경에 일반화하지 않는다. SDK에 있는 기본 사용법은 필요한 비교만 적는다.

## 1. 툴바로 되는 편집이 단축키로는 안 될 수 있다

**관찰 기록.** 한글 입력기가 조합 중일 때 단축키로 액션을 실행하면 `Move*` / `Select*` 계열 액션이 기대대로 캐럿을 움직이지 못하는 문제가 있었다. 같은 기능을 툴바로 클릭하면 문제가 나타나지 않았다는 기록이 있다.

프로젝트에서는 툴바 클릭에 따른 포커스 이탈이 IME 조합을 확정하는 것으로 해석했다. 단축키는 편집창의 포커스를 유지하므로, 두 진입 경로의 입력 상태가 같다고 가정하면 안 된다.

**현재 대응.** 사용자 액션 직전에 `ime::commit_composition()`을 호출한다. 실제 함수 본문의 순서는 다음과 같다.

1. TSF의 `ITfThreadMgr` → 포커스 문서 → 최상위 컨텍스트를 얻어 `TerminateComposition(None)`을 시도한다.
2. 편집창의 부모로 잠시 `SetFocus`했다가 원래 창으로 복귀한다.
3. 여러 HWND를 대상으로 IMM32 조합 확정과 진단을 시도한다.

이 함수는 성공 여부를 반환하지 않으며, 앞 단계가 성공해도 나머지 단계를 실행한다. 오류는 로그에 남기고 계속 진행한다. 따라서 함수를 호출했다는 사실만으로 조합 확정이 보장되지는 않는다.

호출 지점은 두 곳이다. 툴바 등의 사용자 액션은 `HwpUserAction::dispatch()`, 커스텀 단축키는 `shortcut::run_action()`에서 확정을 시도한다. `on_initial_load`, `on_load`, `update_ui`에는 같은 전처리가 자동으로 적용되지 않는다.

근거: [ime.rs](crates/hwp_addon/src/ime.rs)의 `commit_composition`, [hwp_user_action.rs](crates/hwp_addon/src/hwp_user_action.rs)의 `dispatch`, [shortcut.rs](crates/hwp_addon/src/shortcut.rs)의 `run_action`.

### IMM32에서 조합이 안 보여도 입력 중일 수 있다

**관찰 기록.** `ime.rs`에 남아 있는 HWP 2022/2024 진단은 다음과 같다.

- `HwpMainEditWnd`에 대한 `ImmGetContext`가 NULL을 반환했다.
- 부모 창 클래스는 `HwndWrapper[Hwp.exe;;<guid>]` 형태였다.
- 부모 쪽 IMC에서는 `comp_bytes=0`, `open=false`로 나와 IMM32에 조합이 보이지 않았다.

이 결과를 바탕으로 프로젝트는 해당 환경의 입력 처리가 TSF 경로에 있다고 판단했다. IMM32만 반복 호출하는 방법으로는 문제를 해결하지 못했다는 기록이다. HWP의 모든 버전·입력기가 항상 같은 창 구조를 사용한다는 뜻은 아니다.

TSF 진단에서는 `GetFocus()` 같은 호출이 **HRESULT는 S_OK지만 반환 인터페이스 포인터가 NULL**이라 Rust의 `windows` 래퍼에서 오류로 나타나는 경우도 고려한다. 오류 코드만 기록하면 어느 단계에서 객체를 얻지 못했는지 놓치므로 `CoCreateInstance`, `GetFocus`, `GetTop`, 인터페이스 조회, `TerminateComposition`을 나누어 로깅한다.

## 2. 커서 이동 액션을 조합하기보다 위치 기반 선택이 유용했다

**관찰 기록과 구현 선택.** IME 문제 때문에 자동완성을 `MoveWordBegin`이나 `MoveSel*` 액션의 연속 호출만으로 구현하기 어려웠다. 현재는 `GetPos()`로 위치를 얻고 `SelectText()`로 범위를 지정한 뒤 삽입한다.

`HwpEditExt::replace_word_before(prefix, replacement)`는 다음 동작을 묶는다.

1. 현재 문단 내 위치 `pos`를 읽는다.
2. `prefix.chars().count()`만큼 앞의 범위를 선택한다. 문단 시작을 넘으면 0으로 제한한다.
3. 선택 범위를 새 텍스트로 교체한다.

이 헬퍼는 현재 위치에 `prefix`의 내용이 실제로 있는지 비교하지 않는다. 호출자가 문맥을 읽고 올바른 prefix를 넘겨야 한다. 또한 IME 조합 확정을 먼저 시도한 상태에서 사용하도록 설계돼 있다.

근거: [text_edit.rs](crates/hwp_addon/src/text_edit.rs)의 `select_chars_before`, `replace_chars_before`, [IHwpObject 구현](crates/hwp_core/src/ihwpobject/lib.rs)의 `select_text`.

## 3. 스캔 범위와 캐럿 앞 텍스트를 동일시하지 않는다

### `scanEposCurrent`인데 문단 끝까지 읽힌 경우

**관찰 기록.** `read_caret_context()` 주석에는 `ScanEpos::Current`로 종료 위치를 지정해도 `GetText()`가 캐럿을 넘어 문단 끝까지 읽는 경우가 기록돼 있다. 그 결과 캐럿 뒤 텍스트가 prefix 및 next-word 모드 판정에 섞였다.

SDK `HwpAutomation_2504.pdf`의 인쇄 쪽수 22쪽은 `scanEposCurrent`를 캐럿 위치까지의 범위로 설명한다. 여기서 기록할 경험은 그 옵션의 존재가 아니라 **그 설명만으로 캐럿 앞 문자열이 정확히 잘려 나온다고 의존할 수 없었던 사례**다.

**현재 대응.** 문단 전체를 스캔한 뒤 `GetPos()`에서 얻은 문단 내 위치로 직접 앞부분과 뒷글자를 분리한다. 모드 판정 전에 앞부분 끝의 제어 문자도 제거한다. 단, 공백까지 모두 제거하면 next-word 모드를 구분할 수 없으므로 일반적인 `trim()`을 쓰지 않는다.

**미확인 가정.** 현재 구현은 HWP의 문단 내 글자 위치를 Rust `chars()` 인덱스에 대응시킨다. 보충 평면 문자, 이모지, 조합 문자, 문단 안 컨트롤을 포함할 때도 일치하는지는 별도 검증 기록이 없다. 이를 UTF-8 바이트 오프셋으로 바꾸는 것도 근거가 없다.

근거: [dabbrev.rs](crates/hwp_dabbrev/src/dabbrev.rs)의 `read_caret_context`, `expand`.

### 스캔 상태와 반환 텍스트는 함께 처리한다

**구현 선택.** 전체 문서용 `TextWalker`는 `NextParagraph` 상태에 텍스트가 들어 있으면 먼저 `Text`를 내보내고 다음 순회에서 `ParaBreak`를 내보낸다. 상태가 일반 텍스트가 아니라는 이유만으로 반환 문자열을 버리지 않기 위한 처리다.

`EnterControl` 때의 BSTR도 그대로 보존하지만, 이를 반드시 컨트롤 이름이라고 해석하지 않는다. SDK의 `GetText` 설명에는 그 상태의 문자열 의미가 구체적으로 정해져 있지 않다.

`ReleaseScan()` 의무 자체는 SDK에 명시돼 있다. 이 프로젝트의 구현상 교훈은 정상 종료뿐 아니라 중간 `break`, 오류, 이터레이터 폐기에서도 해제되도록 `Drop`에 넣는 것이다. 다만 `read_caret_context()`의 수동 스캔 루프에는 같은 RAII 처리가 적용돼 있지 않으므로 모든 스캔 경로가 동일하게 보호된다고 보면 안 된다.

근거: [text_extract.rs](crates/hwp_addon/src/text_extract.rs)의 `TextWalker::next`, `Pending::ParaBreakDeferred`, `Drop`.

## 4. 자동완성 미리보기는 Undo 이력을 별도로 설계해야 한다

### 후보를 계속 교체하면 Undo가 쌓인다

**관찰 기록.** 후보 A를 B로, B를 C로 단순 교체하던 구현에는 미리보기 과정이 Undo 이력으로 누적되는 문제가 있었다. 이를 수정한 커밋은 `5774765`이다.

**현재 대응.** 새 후보를 표시할 때마다 직전 미리보기를 `undo()`로 되돌린 다음, 확장 전 prefix를 기준으로 새 후보를 삽입한다.

```text
확장 전 상태 → 후보 A 삽입
후보 변경   → Undo → 확장 전 상태 → 후보 B 삽입
확정        → 현재 후보 유지
취소        → Undo → 확장 전 상태
```

목표는 여러 후보를 거쳐도 활성 완성이 Undo 항목 하나로 남게 하는 것이다. 현재 선택과 같은 후보에는 재삽입하지 않으며, 후보가 하나뿐인 경우의 순환도 아무 작업을 하지 않는다.

### next-word 삽입은 직전에 입력한 공백과 합쳐질 수 있다

**관찰 기록.** 공백 뒤에 후보를 단순 삽입하면 Undo 레코드가 직전 입력과 병합되어, 취소 시 후보뿐 아니라 캐럿 앞 공백까지 지워지는 문제가 있었다. `5deef36`에서 빈 범위 선택을 추가했다.

현재 next-word 삽입은 다음 순서다.

```text
GetPos() → SelectText(para, pos, para, pos) → InsertText(candidate)
```

빈 범위 선택이 이 사용 사례에서 Undo 경계를 만드는 우회로로 쓰인다. 최초 삽입과 후보 변경 후 재삽입 양쪽에 적용해야 한다. SDK에 문서화된 범용 트랜잭션 API로 해석하지 않는다.

**확인 한계.** 현재 후보 변경 콜백은 `undo()`와 재삽입 오류를 무시한다. 이 패턴은 각 편집이 성공하고, 중간에 다른 문서 편집이 끼어들지 않는 흐름을 전제로 한다. 오류 상황까지 원상 복구를 보장하는 구현은 아니다.

근거: [dabbrev.rs](crates/hwp_dabbrev/src/dabbrev.rs)의 `insert_candidate`, `run_popup_session`, [ui_popup.rs](crates/hwp_dabbrev/src/ui_popup.rs)의 `cycle`. 이력: `5774765`, `5deef36`.

## 5. 키보드 훅 안에서 팝업 액션을 직접 실행하지 않는다

**관찰 기록과 구현 선택.** `2a36806`에서 단축키 실행을 훅 콜백의 직접 호출에서 전용 메시지 창으로 전달하는 방식으로 바꿨다. 코드 주석은 HWP의 키 처리 도중인 훅 컨텍스트를 벗어나야 창 활성화와 foreground 전환이 정상 동작한다고 설명한다.

현재 흐름은 다음과 같다.

```text
HWP 스레드의 WH_KEYBOARD
  → 단축키 감지, PostMessage(WM_ADDON_ACTION), 해당 키 소비
  → 같은 스레드의 message-only 창 wndproc
  → run_action → IME 조합 확정 시도 → plugin.do_action
```

핵심은 같은 HWP 스레드에서 실행 시점을 늦추는 것이다. 새 작업 스레드로 COM 객체나 팝업을 옮기는 설계가 아니다.

팝업이 열린 동안에는 `set_popup_active(true)`로 단축키 훅을 통과시키고, 팝업이 직접 `Ctrl+/`를 처리한다. 이렇게 해야 후보 순환 키가 다시 새 애드인 액션 호출로 바뀌지 않는다.

**미확인 가정.** 등록 함수 주석에는 최초 1회라고 적혀 있지만, 현재 `register_action_shortcuts()`에는 중복 등록 방지 가드가 없다. `ON_LOAD`에서 다시 호출될 때의 다중 창·재등록 동작을 이미 해결된 것으로 취급하지 않는다.

근거: [shortcut.rs](crates/hwp_addon/src/shortcut.rs)의 `keyboard_hook_proc`, `message_wnd_proc`, `register_action_shortcuts`, [hwp_user_action.rs](crates/hwp_addon/src/hwp_user_action.rs)의 `dispatch`. 이력: `2a36806`.

## 6. 팝업의 메시지·포커스 처리는 문서 편집 결과에도 영향을 준다

### 중첩 메시지 펌프의 종료 범위를 알아야 한다

**구현 선택.** 현재 팝업은 winsafe `WindowMain::run_main()`으로 동기 메시지 펌프를 돌린다. 검토한 winsafe 0.0.29 구현은 창의 `WM_NCDESTROY`에서 `PostQuitMessage`를 호출하고, 중첩된 `run_main` 루프가 그 `WM_QUIT`를 받아 끝난다.

이 정상 종료 흐름을 전제로 HWP의 바깥 메시지 펌프로 복귀한다. 임의의 위치에서 `PostQuitMessage`를 호출해도 안전하다는 뜻은 아니다. 팝업 실행 중 HWP 자체가 종료되는 경우까지 검증됐다는 기록은 없다.

`process_dlg_msgs: false`도 의도가 있는 설정이다. 팝업이 화살표·Enter·Esc를 직접 처리하므로, `IsDialogMessage`가 먼저 처리하는 경로를 끈다. 창 생성 중 panic은 `show()`의 `catch_unwind`에서 취소로 처리하지만, 이것이 애드인 전체 FFI 경계의 panic 보호를 의미하지는 않는다.

### 문자 입력은 키 코드와 다르게 전달한다

**구현 선택.** 현재는 `ToUnicodeEx`를 다시 호출하지 않고, 메시지 펌프의 `TranslateMessage`가 큐에 넣은 `WM_CHAR`를 keydown 처리 중 `PeekMessage`로 꺼낸다. 전달이 필요한 키는 보관했다가 팝업 펌프가 끝난 뒤 HWP에 `KEYDOWN → [CHAR] → KEYUP` 순서로 post한다.

이 순서는 보낸 키를 팝업 펌프가 다시 처리하는 상황을 피하기 위한 것이다. 현재 정책은 Enter 확정, Right/Space 확정 후 전달, Esc 취소, 그 외 키 취소 후 전달이다.

**확인 한계.** `forward_key()`는 KEYUP에도 기존 keydown의 `lParam`을 재사용하고, `drain_char()`는 문자 메시지를 하나만 회수한다. 복잡한 IME 입력, 여러 문자 메시지, 모든 조합키를 일반적인 키보드 재생처럼 처리한다고 보장하지 않는다.

### 창을 닫는 도중의 비활성화와 외부 비활성화를 구분한다

**구현 선택.** 팝업 생성 시 잠깐 발생하는 비활성화로 바로 닫히지 않도록, `got_active`는 활성화 이벤트를 한 번 받은 뒤부터 외부 비활성화를 취소 사유로 인정한다.

또한 Enter 등으로 확정한 뒤 `DestroyWindow`를 호출하면 그 과정의 `WA_INACTIVE`가 확정 결과를 취소로 덮을 수 있다. `explicit_close`를 먼저 세워 이를 막는다. 외부 클릭·Alt+Tab에 의한 비활성화는 취소로 처리한다.

### HWP를 다시 활성화하면서 최대화 상태를 바꾸지 않는다

**관찰 기록.** `ui_popup.rs`에는 `SW_RESTORE`가 최대화 상태를 해제하고 창 위치를 바꾸며, `SW_SHOWNORMAL`도 같은 부작용이 있다는 주석이 있다. 현재는 `BringWindowToTop()`과 `SetForegroundWindow()`를 사용한다.

캐럿 위치는 foreground 스레드의 `GetGUIThreadInfo()`와 `hwndCaret`를 사용해 구하고, 실패하면 화면 중앙을 쓴다. 화면 하단에서 팝업이 잘리는 문제에 대한 위치 조정은 `cdf8f0a`에 남아 있다. 현재의 화면 크기 기반 계산은 모니터별 작업 영역과 DPI 전체를 처리하는 구현은 아니다.

후보 현황 표시도 같은 원칙으로 구현한다. `paint()`에서 실측한 행 높이로 표시 가능한 후보 수를 계산하고, 클릭 판정도 그 값을 공유한다. 상태 표시줄을 추가하면서 그리기 영역만 줄이면 상태 표시줄 클릭이 보이지 않는 후보 선택으로 해석될 수 있다.

근거: [ui_popup.rs](crates/hwp_dabbrev/src/ui_popup.rs)의 `show`, `run_popup`, `drain_char`, `paint`, `try_anchor`, [keyfwd.rs](crates/hwp_addon/src/keyfwd.rs). 이력: `2a36806`, `cdf8f0a`, `376d958`, `589a066`.

## 7. C++ 애드인 인터페이스를 COM 호출 규약으로 일괄 처리하면 안 된다

**구현 근거.** SDK는 C++ 인터페이스와 export 함수 선언을 보여주지만, Rust로 옮길 때의 x86 ABI 구분은 별도로 해석해야 한다. 현재 대상인 `i686-pc-windows-msvc`에서 두 종류를 구분한다.

| 위치 | 현재 Rust 선언 | 의미 |
|---|---|---|
| `QueryUserActionInterface` export | `extern "system"` | SDK의 `__stdcall` export |
| `EnumAction`, `GetActionImage`, `UpdateUI`, `DoAction` | `extern "thiscall"` | C++ 가상 함수. `this`를 ECX로 받음 |

`ffi.rs`에는 vtable을 `extern "C"`나 `extern "system"`으로 선언하면 인자가 깨지고 `this`를 잘못 역참조할 수 있다는 설명이 있다. `IHncUserActionModule`은 이 프로젝트에서 네 개의 C++ 가상 함수로 구성하며, 이름이 인터페이스처럼 보인다고 `IUnknown`의 슬롯을 앞에 넣지 않는다.

문자열도 자료를 대조해야 한다. SDK PDF에는 `LPCTSTR`로 적힌 선언이 있지만, 동봉 C++ 샘플의 `UserActionModule.h`는 액션 이름에 `LPCSTR`를 사용한다. 현재 Rust FFI도 1바이트 C 문자열을 사용한다. GUID와 `DabbrevExpand` 같은 ASCII 식별자가 현재 사용 범위이며, 비ASCII 액션 이름의 인코딩 호환성은 확인되지 않았다.

**미확인 가정.** `EnumAction`은 같은 thread-local 256바이트 버퍼를 재사용한다. HWP가 다음 열거 호출 전에 문자열을 복사한다고 가정하지만, SDK에 그 수명 계약이 명시돼 있지는 않다. 고정 주소라고 해서 앞선 호출의 문자열 내용까지 계속 보존되는 것은 아니다. 현재 이름은 최대 255바이트로 잘린다.

샘플 안에도 별도 확인이 필요한 불일치가 있다. `IsAccessiblePath`는 헤더에서 인자 3개, `.cpp` 정의에서 인자 4개이며, 현재 Rust 매크로는 3개를 export한다. 이 문서에서는 어느 시그니처가 모든 HWP 환경에서 맞는지 결론 내리지 않는다.

근거: [ffi.rs](crates/hwp_addon/src/ffi.rs), [export 매크로](crates/hwp_addon/src/lib.rs), [C++ 헤더](HncUserActionSample/HwpUserAction/UserActionModule.h), [C++ 구현](HncUserActionSample/HwpUserAction/UserActionModule.cpp). SDK 비교: `Addon-Action.pdf` 인쇄 쪽수 4–5쪽.

## 8. 빌려 받은 COM 포인터와 소유한 COM 참조를 구분한다

**샘플과 구현에서 도출한 규칙.** 애드인 콜백으로 들어온 `pObject`는 HWP가 소유한 포인터다. C++ 샘플은 `AttachDispatch(pObject, FALSE)`로 잠시 감싸고 `DetachDispatch()`한다.

현재 Rust의 `HwpObject::from_raw_dispatch()`도 참조를 새로 얻은 것으로 취급하지 않는다. 임시 래퍼가 Drop되면서 HWP가 소유한 참조를 Release하지 않도록 호출 측에서 `mem::forget()`을 해야 한다. 일반적인 OLE 클라이언트의 `CoCreateInstance` 반환 객체와 수명 관리가 다르다.

반면 `HwpObject::clone()`은 `IDispatch`를 Clone하여 AddRef로 자신의 참조를 얻는다. 팝업의 `'static` 클로저에는 이 소유한 복제본을 넘기며, Drop 시 자신의 참조를 Release한다. 이는 새 HWP 인스턴스나 새 문서를 만드는 동작이 아니다.

**현재 구현의 한계.** `shortcut::run_action()`에서는 콜백 호출의 `?`가 마지막 `mem::forget()`보다 앞에 있다. 콜백이 오류를 반환하면 정상 경로의 수명 처리까지 도달하지 않는다. 수동 `forget()` 패턴을 복사할 때는 조기 반환과 panic 경로도 함께 검토해야 한다. 여기서는 문제를 기록하며 코드는 변경하지 않았다.

근거: [hwp_obj.rs](crates/hwp_core/src/hwp_obj.rs)의 `from_raw_dispatch`, `Clone`, [ffi.rs](crates/hwp_addon/src/ffi.rs)의 `with_hwp_object`, [shortcut.rs](crates/hwp_addon/src/shortcut.rs)의 `run_action`, [C++ 구현](HncUserActionSample/HwpUserAction/UserActionModule.cpp)의 `DoAction`.

### `Version`의 타입을 한 가지로 고정하지 않는다

**구현에 남은 호환성 기록.** `detect_version()` 주석은 OLE 클라이언트에서 `"10, 0, 0, 14727"` 같은 문자열, 애드인에서 정수 코드가 오는 경우를 구분한다. 현재는 BSTR 변환을 먼저 시도하고, 그다음 i32 변환을 시도한다.

SDK의 `Version` 설명은 바이트별 정수 구조를 제시한다. 프로젝트의 추가 지식은 문자열 경로도 다루도록 구현돼 있다는 점이다. 정확한 환경별 반환 타입 표는 없으며, 변환 실패 시 `Other(0, 0)`으로 처리된다. `V2018`, `V2022`, `V2024` enum이 있다는 사실은 해당 버전들의 모든 기능을 실환경 검증했다는 뜻이 아니다.

근거: [hwp_obj.rs](crates/hwp_core/src/hwp_obj.rs)의 `detect_version`, [hwp_ver.rs](crates/hwp_core/src/hwp_ver.rs). SDK 비교: `Version(Property)`.

## 9. 문서 내용과 UI의 수명은 플러그인 수명과 다르다

### 단어 캐시 제거는 HWP 규칙이 아니라 프로젝트의 설계 결정이다

**구현 선택.** 과거 `hwp_dabbrev`는 문서 경로를 키로 단어·전체 텍스트 캐시를 유지했다. 저장하지 않은 문서를 빈 문자열 키로 합치고, 후보 소진 뒤 다시 추출하는 단계도 있었다.

`d8b0bb5`에서 구조 단순화를 위해 이 상태를 제거했다. 현재는 액션마다 문서 전체에서 단어를 한 번 추출하고, 모든 후보를 만든 뒤 팝업에 넘긴다. 후보를 순환하는 동안에는 재추출하지 않는다. 캐시 무효화나 선택 이력을 유지하지 않으며, 그 대가로 매 액션마다 전체 스캔 비용이 든다.

이는 과거의 캐시 방식이 HWP에서 반드시 불가능하다는 결론이 아니다. 다시 캐시를 도입한다면 저장 전 문서 구분, 편집 반영, 문서 전환, 저장 경로 변경을 별도로 설계해야 한다는 경험이다.

단어의 정의도 프로젝트 정책이다. 현재는 영숫자와 내부 `_`, `-`, `:`를 허용하고, 선두 연결 문자는 제거한다. 마침표 `.`는 `92d8147` 이후 구분자다. 이 규칙은 한글 내장 단어 이동 액션의 규칙과 동일하다고 확인된 것이 아니다.

전체 텍스트를 단어 배열로 평탄화하면서 컨트롤 경계는 별도로 유지하지 않는다. next-word는 그 배열의 인접 단어를 기준으로 하므로, 화면상 같은 문장에 속한 관계만 보장하는 방식은 아니다.

근거: [dabbrev.rs](crates/hwp_dabbrev/src/dabbrev.rs), [lib.rs](crates/hwp_dabbrev/src/lib.rs)의 `extract_all_words`, `is_word_char`, `strip_leading_nonword`. 이력: `5deef36`, `92d8147`, `d8b0bb5`.

### 툴바 변경이 안 보이면 직렬화 경로 가드도 확인한다

**구현에서의 추론.** `setup_toolbar()`는 `ChangeSerializePath()` 후 `IsNewSerializePath()`가 false이면 UI 생성을 생략한다. 따라서 DLL에서 버튼 정의를 바꿨는데 화면이 그대로일 때, 새 바이너리 로드 여부와 함께 이 조기 반환 여부를 확인할 필요가 있다.

현재 구성은 `serialize_path`를 리본 탭 UID로도 사용한다. 개발 중 이 값을 바꾸는 일은 단순한 표시 이름 변경과 같지 않다. 기존 사용자 UI 설정의 갱신·마이그레이션 방식은 별도 설계 대상이다.

근거: [hwp_user_action.rs](crates/hwp_addon/src/hwp_user_action.rs)의 `ToolbarConfig`, `setup_toolbar`. 직렬화 API 자체는 SDK에 있는 내용이다.

## 10. 진단 결과를 읽을 때 알아둘 점

로그 기본 경로는 Windows의 `%LOCALAPPDATA%\HwpAddon\hwp_addon_debug.log`다. `ime`, `DoAction`, `dabbrev`, `ui_popup`, `com_util` 등의 태그로 실패 단계와 호출 흐름을 구분한다.

**구현상 주의.** `hwp_core`와 `hwp_addon`의 로그 모듈은 각각 독립적인 `OnceLock`을 사용하지만 같은 파일을 `truncate(true)`로 연다. 한쪽이 나중에 처음 로그를 남기면 앞쪽 로그가 지워질 수 있다. 이 때문에 초기 진단 기록이 없다는 사실만으로 해당 코드가 실행되지 않았다고 단정하지 않는다. 파일을 열지 못해도 로깅은 조용히 생략된다.

WSL 빌드에서는 Rust 컴파일과 Windows 링커 실행을 구분한다. 현재 `.cargo/config.toml`의 링커는 `../msvc-linker/linker-x86.sh`다. 이 작업 세션에서는 링커가 저장소 밖에 보조 파일을 쓰는 단계의 읽기 전용 오류와 WSL 소켓 오류가 있었고, 실행 권한 범위를 조정한 뒤 같은 빌드가 성공했다. 이는 HWP API 호출 오류와는 다른 진단 범주다.

기준 커밋의 의존성과 팝업 변경은 `cargo build -p hwp_dabbrev --locked`로 32비트 Windows DLL 빌드를 확인했다. 빌드 성공은 IME, 포커스, Undo, 문서별 동작의 실환경 검증을 대신하지 않는다.

근거: [애드인 로그](crates/hwp_addon/src/debug.rs), [코어 로그](crates/hwp_core/src/debug.rs), [링커 설정](.cargo/config.toml).

## 기존 문서·주석을 읽을 때의 정정 사항

오래된 설명을 다시 구현 근거로 삼지 않도록 차이를 남긴다. 아래 항목은 현재 코드와 SDK를 대조한 결과이며, 이 문서 작성으로 해당 소스나 과거 문서를 일괄 수정하지는 않았다.

| 자료 | 남아 있는 설명 | 현재 기준 |
|---|---|---|
| [C++ 샘플 분석](HncUserActionSample_분석.md) §9.2 | vtable도 `extern "system"` | x86 C++ 가상 함수는 `extern "thiscall"`. export와 구분 |
| [단축키 구조 문서](crates/hwp_addon/shortcut.md) | 훅에서 `run_action` 직접 실행 | 현재는 `PostMessage` → 메시지 창의 wndproc → `run_action` |
| [ime.rs](crates/hwp_addon/src/ime.rs) 상단 시도 순서 | SetFocus → TSF → IMM32 | 실제 함수 본문은 TSF → SetFocus → IMM32 |
| [text_edit.rs](crates/hwp_addon/src/text_edit.rs) 모듈 설명 | `on_load` / `update_ui`에서도 조합 확정 상태를 기대하는 표현 | 자동 확정 시도는 사용자 액션의 `dispatch`와 `run_action` 경로에 있음 |
| [hwp_com 설명 주석](crates/hwp_com/src/lib.rs) | ROT 이름 `!IHwpObject` | 실제 필터는 `!HwpObject`. SDK 예제도 `!HwpObject.130` |

## 새 관찰을 추가할 때

한 항목에 **증상 → 재현 조건 → 진단 결과 → 대응 → 확인 범위 → 코드·커밋 근거**를 남긴다. 특히 HWP/Windows/IME 버전과 툴바·단축키 중 어느 경로였는지 기록하면, 버전 차이와 입력 상태 차이를 구분하기 쉽다.

우회 방법은 성공 경로뿐 아니라 원래 입력 보존까지 확인한다. 이 프로젝트에서 의미 있는 확인 사례는 조합 중 단축키와 툴바 비교, 문단 중간 캐럿의 앞뒤 텍스트, 여러 후보를 거친 뒤 취소, 공백 뒤 next-word 취소, 최대화된 HWP로 복귀, 저장하지 않은 여러 문서 사이의 전환이다. 이 목록은 이후 재현을 위한 항목이며, 모든 조합을 이미 검증했다는 기록은 아니다.
