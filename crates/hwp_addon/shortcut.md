# Shortcut 구조와 콜스택

## 개요

HWP 애드인에서 커스텀 단축키를 지원하기 위한 모듈. Windows Keyboard Hook(`WH_KEYBOARD`)을 HWP 스레드에 설치하여 키 입력을 가로채고, 등록된 단축키와 매칭되면 플러그인의 `do_action`을 직접 호출한다.

## 핵심 구조체/트레잇

```
ShortcutKey          — ActionMeta에 넣는 단축키 선언 (Modifiers + VIRTUAL_KEY)
Modifiers            — Alt/Ctrl/Shift 조합
ShortCut trait       — 사용자가 직접 구현할 수도 있는 단축키 트레잇 (현재 메인 흐름 미사용)
ActionCallback       — type-erased 플러그인 포인터 + do_action 함수 포인터
```

## 전역 상태

| static                  | 타입                              | 용도                                 |
|-------------------------|-----------------------------------|--------------------------------------|
| `ACTIONSWITHSHORTCUT`   | `Mutex<Vec<&'static ActionMeta>>` | 단축키가 있는 액션 메타 목록         |
| `HWP_DISPATCH`          | `AtomicPtr<c_void>`               | 현재 HWP IDispatch raw 포인터        |
| `ACTION_CALLBACK`       | `Mutex<Option<ActionCallback>>`   | 플러그인의 do_action을 호출하는 콜백 |

## 콜스택

### 1. 초기화 (DLL 로드 → 훅 설치)

```
HWP.exe
 └─ DLL 로드 → QueryUserActionInterface()
     └─ HWP가 EnumAction(n) 반복 호출 → 액션 목록 수집
     └─ HWP가 DoAction("{ON_INITIAL_LOAD GUID}", pObj) 호출
         └─ tramp_do_action<T>()                    [ffi.rs]
             └─ plugin.dispatch("ON_INITIAL_LOAD", hwp)
                 └─ self.on_initial_load(hwp)        → 최초 1회 초기화
     └─ HWP가 DoAction("{ON_LOAD GUID}", pObj) 호출
         └─ tramp_do_action<T>()                    [ffi.rs]
             └─ extract_plugin<T>(this)
             └─ with_hwp_object(pObj, |hwp| ...)
                 └─ plugin.dispatch("ON_LOAD", hwp) [hwp_user_action.rs]
                     ├─ set_hwp_dispatch(hwp)        → HWP_DISPATCH에 저장
                     ├─ set_action_callback(self)     → ACTION_CALLBACK에 저장
                     ├─ register_action_shortcuts(self)  [shortcut.rs]
                     │   ├─ actions()에서 shortcut이 있는 항목마다
                     │   │   └─ ACTIONSWITHSHORTCUT에 &ActionMeta push
                     │   └─ ACTIONSWITHSHORTCUT가 비어있지 않으면:
                     │       └─ SetWindowsHookExW(WH_KEYBOARD, keyboard_hook_proc, HWP스레드)
                     │           (훅 핸들은 저장하지 않음)
                     └─ self.on_load(hwp)             → setup_toolbar 등
```

### 2. 단축키 실행 (키 입력 → 액션 실행)

```
HWP 메시지 루프 (HWP 스레드)
 └─ WM_KEYDOWN 수신
     └─ keyboard_hook_proc(nCode, wParam, lParam)   [shortcut.rs:209]
         ├─ lParam bit31 == 0 (키 누름) 확인
         ├─ vk = VIRTUAL_KEY(wParam)
         ├─ current_mods = Modifiers::current()      ← GetAsyncKeyState로 조회
         ├─ ACTIONSWITHSHORTCUT.lock() → 순회
         │   └─ sc.shortcut.key == vk && sc.shortcut.modifiers == current_mods 이면:
         │       └─ run_action(sc.name)              [shortcut.rs:154]
         │           ├─ HWP_DISPATCH.load() → null이면 조기 반환
         │           ├─ ime::commit_composition()     ★ IME 조합 확정
         │           │   ├─ terminate_tsf_composition()  — TSF 경로 (핵심)
         │           │   ├─ try_setfocus_bounce()        — SetFocus 해킹
         │           │   └─ try_imm32_all_candidates()   — IMM32 (진단용)
         │           ├─ HwpObject::from_raw_dispatch(raw)
         │           ├─ ACTION_CALLBACK.lock()
         │           │   └─ (cb.call_fn)(cb.plugin_ptr, action_name, &hwp)
         │           │       └─ plugin.do_action(name, hwp)  ← 사용자 구현
         │           └─ std::mem::forget(hwp)         ← Release 방지
         └─ return LRESULT(1)                        ← keystroke 소비
     (매칭 안 되면)
     └─ CallNextHookEx() → HWP 기본 처리
```

### 3. 툴바 클릭 경로 (비교)

```
HWP.exe → 사용자가 툴바/리본 버튼 클릭
 └─ DoAction("ActionName", pObj)
     └─ tramp_do_action<T>()                        [ffi.rs]
         └─ plugin.dispatch(action_name, hwp)       [hwp_user_action.rs]
             └─ ime::commit_composition()            ★
             └─ self.do_action(action_name, hwp)     ← 사용자 구현
```

## 단축키 경로 vs 툴바 경로 차이

| 구분          | 단축키 경로                         | 툴바 경로                            |
|---------------|-------------------------------------|--------------------------------------|
| 진입점        | `keyboard_hook_proc` → `run_action` | HWP → `tramp_do_action`              |
| dispatch 경유 | **아니오** — do_action 직접 호출    | **예** — dispatch() 거침             |
| IME commit    | `run_action` 내부에서 명시적 호출   | `dispatch` 내부 + focus-out 부수효과 |
| 키 소비       | `LRESULT(1)` 반환으로 HWP 차단      | 해당 없음                            |

## 사용 예 (hwp_dabbrev)

```rust
fn actions(&self) -> &'static [ActionMeta] {
    &[ActionMeta {
        name: "DabbrevExpand",
        label: "dabbrev",
        image_index: 0,
        shortcut: Some(ShortcutKey {
            modifiers: Modifiers { alt: false, ctrl: true, shift: false },
            key: VK_OEM_2,  // Ctrl+/
        }),
    }]
}
```

`ActionMeta`에 `shortcut`을 지정하면 `ON_LOAD` 시 자동으로 키보드 훅이 설치된다. `do_action`만 구현하면 툴바 클릭과 단축키 양쪽 모두에서 호출된다.
