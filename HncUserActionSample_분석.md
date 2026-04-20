# HncUserActionSample 분석

HWP 애드인(UserAction) DLL 개발을 위한 C++/MFC 샘플 프로젝트.

---

## 1. 프로젝트 구조

```
HncUserActionSample/
├── HncUserAction.sln                    # Visual Studio 솔루션
├── Image/                               # 툴바 비트맵 리소스
│   ├── toolbar_normal.bmp               # 소형 툴바 아이콘 (활성)
│   ├── toolbar_inactive.bmp             # 소형 툴바 아이콘 (비활성)
│   ├── toolbar_Large.bmp                # 대형 툴박스 아이콘 (활성)
│   └── toolbar_Large_Inactive.bmp       # 대형 툴박스 아이콘 (비활성)
├── HwpUserAction/                       # 메인 DLL 소스
│   ├── HwpUserAction.h/.cpp             # MFC DLL 앱 클래스 (CHwpUserActionApp)
│   ├── HwpUserAction.def                # DLL 익스포트 정의
│   ├── UserActionModule.h/.cpp          # 핵심: IHncUserActionModule 인터페이스 구현
│   ├── UserActions.h/.cpp               # 실제 액션 핸들러 함수들
│   ├── UserActionMap.h                  # 액션맵 매크로 + 타입 정의
│   ├── UserButtonMap.h                  # (비어있음) 버튼맵 정의용
│   ├── UserActionsUUIDSTR.h             # GUID 상수 + 시리얼라이즈 경로
│   ├── HwpObjects.h                     # COM 래퍼 클래스 헤더 일괄 include
│   ├── stdafx.h                         # PCH: MFC + OLE + 전체 헤더 include
│   └── resource.h                       # 비트맵 리소스 ID (IDB_TOOLBAR_*)
├── Objects/
│   ├── HwpObjects/                      # IHwpObject 등 HWP COM 객체 래퍼 (~150개)
│   └── HncObjects/                      # 툴박스 UI COM 객체 래퍼 (9개)
```

## 2. DLL 진입점 — 익스포트 함수

`HwpUserAction.def`에서 2개 함수를 익스포트:

### `QueryUserActionInterface() -> IHncUserActionModule*`

HWP가 DLL 로드 시 호출하는 **핵심 진입점**. `static CUserActionModule` 싱글턴을 리턴.

```cpp
// UserActionModule.cpp:135
IHncUserActionModule* __stdcall QueryUserActionInterface()
{
    static CUserActionModule uam;
    return static_cast<IHncUserActionModule*>(&uam);
}
```

### `IsAccessiblePath(HWND, LONG, LPCTSTR, LPCTSTR) -> BOOL`

HWP가 파일 접근 전에 호출하는 보안 콜백. 샘플에서는 항상 `TRUE` 리턴.

## 3. IHncUserActionModule — 핵심 인터페이스

HWP가 호출하는 vtable 인터페이스. **4개 가상 메서드**로 구성:

```cpp
interface IHncUserActionModule
{
    // 등록된 액션 이름을 순회 (iterator 패턴)
    virtual LPCSTR EnumAction(int nIterator) = 0;

    // 액션에 대응하는 툴바 비트맵 + 이미지 인덱스 반환
    virtual BOOL GetActionImage(LPCSTR szAction, UINT uState,
                                HBITMAP* lphBitmap, int* lpnImageIndex) = 0;

    // 액션의 UI 상태 반환 (0=보통, DFCS_INACTIVE, DFCS_CHECKED, DFCS_BUTTON3STATE)
    virtual BOOL UpdateUI(LPCSTR szAction, LPDISPATCH pObject, UINT* lpuState) = 0;

    // 액션 실행 — pObject는 IHwpObject (CHwpObject로 래핑)
    virtual int DoAction(LPCSTR szAction, LPDISPATCH pObject) = 0;
};
```

### vtable 레이아웃 (C ABI 관점)

| 인덱스 | 메서드 | 시그니처 |
|--------|--------|----------|
| 0 | `EnumAction` | `(this, i32) -> *const c_char` |
| 1 | `GetActionImage` | `(this, *const c_char, u32, *mut HBITMAP, *mut i32) -> BOOL` |
| 2 | `UpdateUI` | `(this, *const c_char, LPDISPATCH, *mut u32) -> BOOL` |
| 3 | `DoAction` | `(this, *const c_char, LPDISPATCH) -> i32` |

## 4. 액션 등록 시스템

### 4.1. USERACTION 구조체

```cpp
typedef BOOL(*PFNUSERACTION)(CHwpObject& pObject);
typedef UINT(*PFNUSERACTION_UPDATEUI)(CHwpObject& pObject);

typedef struct tagUserAction {
    LPCSTR szActionName;                    // GUID 문자열 (액션 식별자)
    int nImageIndex;                        // 비트맵 내 아이콘 인덱스
    PFNUSERACTION pfnUserAction;            // 실행 함수 포인터
    PFNUSERACTION_UPDATEUI pfnUserActionUpdateUI;  // UI 상태 함수 포인터 (NULL 가능)
} USERACTION;
```

### 4.2. 액션맵 매크로

```cpp
#define BEGIN_USERACTION_MAP()  USERACTION g_arUserActions[] = {
#define END_USERACTION_MAP()    };
#define USERACTION_ENTRY(NAME, IMAGEINDEX, ACTIONFP, UPDATEUIFP) {NAME, IMAGEINDEX, ACTIONFP, UPDATEUIFP},
```

### 4.3. 샘플의 액션맵 (UserActionMap.h)

```cpp
BEGIN_USERACTION_MAP()
    // 예약된 시스템 액션 (HWP가 자동 호출)
    USERACTION_ENTRY(UUIDSTR_ON_INITIAL_LOAD,        0,  OnInitialLoad,  0)
    USERACTION_ENTRY(UUIDSTR_ON_LOAD,                0,  OnLoad,         0)

    // 사용자 정의 액션 (툴바/툴박스 버튼에 바인딩)
    USERACTION_ENTRY(UUIDSTR_ON_USERACTION_TOOLBAR,  64, OnUserAction1,  0)
    USERACTION_ENTRY(UUIDSTR_ON_USERACTION_TOOLBOX,  64, OnUserAction1,  0)
END_USERACTION_MAP()
```

### 4.4. 예약 GUID 상수

```cpp
// HWP 내부 예약 GUID — 변경 불가
#define UUIDSTR_ON_INITIAL_LOAD     "{B91A2981-A001-44a9-933F-5BF70A747967}"  // DLL 최초 로드
#define UUIDSTR_ON_LOAD             "{3E4DC866-051C-4989-820E-DADC1E6264B9}"  // 문서 로드마다

// 사용자 정의 GUID — 다른 GUID와 겹치면 안됨
#define UUIDSTR_ON_USERACTION_TOOLBAR  "{752C9DDE-6AB6-40EA-8C3F-DB18D2779482}"
#define UUIDSTR_ON_USERACTION_TOOLBOX  "{64B5DAFA-F5E2-4035-9ACF-522946227C80}"

// 레지스트리 시리얼라이즈 경로
#define USERACTION_SERIALIZE_PATH   "HwpUserAction"

// 툴박스 UI 요소 UID
#define USERTOOLBOXTAB_UID      "{EE1DAE22-7B0F-41ED-9CF7-D657A5BE0841}"
#define USERTOOLBOX_UID         "{189D2B2A-60DC-4C09-B234-971141BB45ED}"
#define USERTOOLBOXGROUP_UID    "{6E999844-C8AD-4FC1-B993-A742276EC893}"
```

## 5. 액션 라이프사이클

### 5.1. DLL 로드 흐름

```
HWP 시작
  → DLL 로드 (MFC DllMain → CHwpUserActionApp::InitInstance)
  → QueryUserActionInterface() 호출 → IHncUserActionModule* 획득
  → EnumAction(0..N) 반복 → 등록된 액션 이름 수집
  → DoAction(UUIDSTR_ON_INITIAL_LOAD, pHwpObject) 호출
  → DoAction(UUIDSTR_ON_LOAD, pHwpObject) 호출 (문서 열릴 때마다)
```

### 5.2. OnInitialLoad — 초기화

샘플에서는 비어있음 (`return TRUE`). 최초 1회 호출.

### 5.3. OnLoad — UI 생성

문서가 열릴 때마다 호출. 툴바와 툴박스에 버튼을 등록:

```
HwpObject
  → XHwpWindows → Active XHwpWindow → XHwpToolbarLayout
      ├── ChangeSerializePath("HwpUserAction")
      ├── IsNewSerializePath()  ← 이미 등록됐으면 스킵
      │
      ├── [툴바 생성]
      │   CreateToolbar("사용자 정의", 0)
      │   └── CreateToolbarButton("툴바버튼", UUIDSTR_TOOLBAR, 0x03)
      │       └── InsertToolbarButton(button, -1)
      │
      └── [툴박스 생성] (GetToolBoxToolBar)
          InsertToolBoxTab(-1, UID, "추가 기능")
          └── InsertToolBox(-1, UID, "멀티미디어")
              └── GetLayout(0)
                  └── InsertGroup(-1, UID, LARGEICON, rows=1, cols=4)
                      └── CreateToolBoxItemButtonEx("툴박스버튼", UUIDSTR_TOOLBOX, STDTB_BTN, 0x03)
                          └── InsertItem(-1, pDispatch)
```

### 5.4. DoAction 디스패치

```cpp
// UserActionModule.cpp:110
BOOL CUserActionModule::DoAction(LPCSTR szAction, LPDISPATCH pObject)
{
    for (i = 0; i < NELEM(g_arUserActions); i++) {
        if (!strcmp(g_arUserActions[i].szActionName, szAction)) {
            CHwpObject app;
            app.AttachDispatch(pObject, FALSE);  // IDispatch → CHwpObject 래핑
            BOOL res = g_arUserActions[i].pfnUserAction(app);
            app.DetachDispatch();                // Release 없이 분리
            return res;
        }
    }
    return FALSE;
}
```

핵심 패턴: `LPDISPATCH`를 `CHwpObject`에 `AttachDispatch(pObject, FALSE)`로 래핑 → 함수 호출 → `DetachDispatch()`로 분리. `FALSE`는 Release를 호출하지 않겠다는 의미 (소유권을 가져가지 않음).

### 5.5. GetActionImage — 이미지 디스패치

액션 이름과 상태(`DFCS_INACTIVE` 여부)에 따라 적절한 비트맵 핸들 + 인덱스 반환. 툴바용(small)과 툴박스용(large) 비트맵을 구분.

### 5.6. UpdateUI — UI 상태 갱신

등록된 `pfnUserActionUpdateUI` 콜백이 있으면 호출하여 UI 상태 반환. 없으면 `*lpuState = 0` (보통상태).

| 반환값              | 의미          |
|---------------------|---------------|
| `0`                 | 보통 상태     |
| `DFCS_INACTIVE`     | 비활성 (회색) |
| `DFCS_CHECKED`      | 체크됨 (토글) |
| `DFCS_BUTTON3STATE` | 중간 상태     |

## 6. COM 객체 래퍼 구조

### 6.1. 래핑 패턴

모든 COM 객체 래퍼는 MFC의 `COleDispatchDriver`를 상속. `InvokeHelper()`로 IDispatch 호출:

```cpp
class CHwpObject : public COleDispatchDriver
{
    // property get: InvokeHelper(DISPID, DISPATCH_PROPERTYGET, VT_*, &result, NULL)
    // property put: InvokeHelper(DISPID, DISPATCH_PROPERTYPUT, VT_EMPTY, NULL, parms, value)
    // method:       InvokeHelper(DISPID, DISPATCH_METHOD, VT_*, &result, parms, args...)
};
```

### 6.2. HwpObjects — HWP COM 객체 래퍼

| 클래스                     | COM 인터페이스        | 주요 역할                                        |
|----------------------------|-----------------------|--------------------------------------------------|
| `CHwpObject`               | `IHwpObject`          | **메인 API 진입점**. 문서 조작의 모든 것         |
| `CDHwpAction`              | `IDHwpAction`         | 액션 생성/실행 (CreateSet, GetDefault, Execute)  |
| `CDHwpParameterSet`        | `IDHwpParameterSet`   | 액션 파라미터 셋 (Item, SetItem, CreateItemSet)  |
| `CDHwpParameterArray`      | `IDHwpParameterArray` | 파라미터 배열                                    |
| `CDHwpCtrlCode`            | `IDHwpCtrlCode`       | 문서 컨트롤 코드 (CtrlID, Properties, Next/Prev) |
| `CHwpObjectEvents`         | `IHwpObjectEvents`    | 이벤트 핸들러 (Quit, NewDocument, BeforeSave 등) |
| `CXHwpWindows`             | `IXHwpWindows`        | 윈도우 컬렉션                                    |
| `CXHwpWindow`              | `IXHwpWindow`         | 개별 윈도우 (위치, 크기, Visible, ToolbarLayout) |
| `CXHwpToolbarLayout`       | `IXHwpToolbarLayout`  | 툴바 레이아웃 관리자                             |
| `CXHwpToolbar`             | `IXHwpToolbar`        | 개별 툴바                                        |
| `CXHwpToolbarButton`       | `IXHwpToolbarButton`  | 툴바 버튼                                        |
| 기타 `CH*` 클래스 (~130개) | 각종 ParameterSet     | 액션별 파라미터 정의 래퍼                        |

### 6.3. HncObjects — 툴박스 UI 래퍼

| 클래스                     | 역할                                 |
|----------------------------|--------------------------------------|
| `CXHncToolBoxToolbar`      | 툴박스 최상위. 탭 생성, 버튼 생성    |
| `CXHncToolBoxTab`          | 툴박스 탭 (UID, Name, InsertToolBox) |
| `CXHncToolBox`             | 툴박스 컨테이너 (GetLayout)          |
| `CXHncToolBoxLayout`       | 레이아웃 (InsertGroup)               |
| `CXHncToolBoxGroup`        | 그룹 (InsertItem)                    |
| `CXHncToolBoxItem`         | 아이템 기본                          |
| `CXHncToolBoxItemButton`   | 버튼 아이템                          |
| `CXHncToolBoxItemMenu`     | 메뉴 아이템                          |
| `CXHncToolBoxItemComboBox` | 콤보박스 아이템                      |

## 7. CHwpObject(IHwpObject) 주요 API

### Properties (DISPID 0x01~0x0F)

| DISPID | Property          | 타입                | 설명               |
|--------|-------------------|---------------------|--------------------|
| 0x01   | `IsModified`      | BOOL (get)          | 문서 수정 여부     |
| 0x02   | `IsEmpty`         | BOOL (get)          | 빈 문서 여부       |
| 0x03   | `EditMode`        | long (get/put)      | 편집 모드          |
| 0x04   | `SelectionMode`   | long (get)          | 선택 모드          |
| 0x05   | `CurFieldState`   | long (get)          | 현재 필드 상태     |
| 0x06   | `PageCount`       | long (get)          | 페이지 수          |
| 0x07   | `CellShape`       | IDispatch (get/put) | 셀 모양            |
| 0x08   | `CharShape`       | IDispatch (get/put) | 글자 모양          |
| 0x09   | `HeadCtrl`        | IDispatch (get)     | 첫 번째 컨트롤     |
| 0x0A   | `LastCtrl`        | IDispatch (get)     | 마지막 컨트롤      |
| 0x0B   | `CurSelectedCtrl` | IDispatch (get)     | 현재 선택된 컨트롤 |
| 0x0C   | `ParaShape`       | IDispatch (get/put) | 문단 모양          |
| 0x0D   | `ParentCtrl`      | IDispatch (get)     | 부모 컨트롤        |
| 0x0E   | `ViewProperties`  | IDispatch (get/put) | 보기 속성          |
| 0x0F   | `Path`            | BSTR (get)          | 문서 경로          |

### Sub-objects (DISPID 0x4E20~)

| DISPID | Property         | 타입      |
|--------|------------------|-----------|
| 0x4E20 | `Application`    | IDispatch |
| 0x4E21 | `XHwpMessageBox` | IDispatch |
| 0x4E22 | `XHwpDocuments`  | IDispatch |
| 0x4E23 | `XHwpWindows`    | IDispatch |
| 0x4E24 | `HParameterSet`  | IDispatch |
| 0x4E25 | `HAction`        | IDispatch |
| 0x4E26 | `XHwpODBC`       | IDispatch |
| 0x4E85 | `Version`        | BSTR      |

### 문서 조작 메서드 (DISPID 0x2710~)

| DISPID | 메서드            | 시그니처                                         | 설명                    |
|--------|-------------------|--------------------------------------------------|-------------------------|
| 0x2710 | `Open`            | (filename, Format, arg) → BOOL                   | 문서 열기               |
| 0x2711 | `Save`            | (save_if_dirty) → BOOL                           | 저장                    |
| 0x2712 | `SaveAs`          | (Path, Format, arg) → BOOL                       | 다른 이름 저장          |
| 0x2713 | `Insert`          | (Path, Format, arg)                              | 삽입                    |
| 0x2714 | `SelectText`      | (spara, spos, epara, epos) → BOOL                | 텍스트 선택             |
| 0x2715 | `CreateField`     | (Direction, memo, name) → BOOL                   | 필드 생성               |
| 0x2716 | `MoveToField`     | (Field, Text, start, select) → BOOL              | 필드로 이동             |
| 0x2717 | `FieldExist`      | (Field) → BOOL                                   | 필드 존재 확인          |
| 0x2718 | `GetFieldText`    | (Field) → BSTR                                   | 필드 텍스트 조회        |
| 0x2719 | `PutFieldText`    | (Field, Text)                                    | 필드 텍스트 설정        |
| 0x271A | `RenameField`     | (old, new)                                       | 필드 이름 변경          |
| 0x271B | `GetCurFieldName` | (option) → BSTR                                  | 현재 필드명             |
| 0x2720 | `MovePos`         | (moveID, Para, pos) → BOOL                       | 커서 이동               |
| 0x2721 | `InitScan`        | (option, Range, spara, spos, epara, epos) → BOOL | 스캔 초기화             |
| 0x2722 | `ReleaseScan`     | ()                                               | 스캔 해제               |
| 0x2723 | `GetText`         | (&Text) → long                                   | 텍스트 읽기             |
| 0x2724 | `GetPos`          | (&List, &Para, &pos)                             | 현재 위치               |
| 0x2725 | `SetPos`          | (List, Para, pos) → BOOL                         | 위치 설정               |
| 0x2727 | `GetTextFile`     | (Format, option) → VARIANT                       | 텍스트 파일 형태로 얻기 |
| 0x2728 | `SetTextFile`     | (data, Format, option) → long                    | 텍스트 파일 형태로 설정 |
| 0x272A | `Run`             | (ActID)                                          | 액션 실행               |
| 0x272B | `LockCommand`     | (ActID, isLock)                                  | 명령 잠금               |
| 0x272D | `InsertPicture`   | (Path, ...) → IDispatch                          | 그림 삽입               |
| 0x272F | `CreateAction`    | (actidstr) → IDispatch                           | 액션 객체 생성          |
| 0x2730 | `InsertCtrl`      | (CtrlID, initparam) → IDispatch                  | 컨트롤 삽입             |
| 0x2734 | `RegisterModule`  | (ModuleType, ModuleData) → BOOL                  | 모듈 등록               |
| 0x2735 | `ReplaceAction`   | (OldActionID, NewActionID) → BOOL                | 액션 교체               |

### 유틸리티 메서드 (DISPID 0x7530~)

| DISPID  | 메서드                             | 설명                            |
|---------|------------------------------------|---------------------------------|
| 0x7530  | `Quit`                             | HWP 종료                        |
| 0x7535  | `MiliToHwpUnit`                    | mm → HWP 단위 변환              |
| 0x7536  | `PointToHwpUnit`                   | pt → HWP 단위 변환              |
| 0x7537  | `RGBColor`                         | RGB → HWP 색상값                |
| 0x7538~ | `HwpLineWidth`, `HwpLineType`, ... | 문자열 → 열거값 변환 유틸리티들 |
| 0x758C  | `IsActionEnable`                   | 액션 활성화 여부                |
| 0x7590  | `GetPageText`                      | 페이지 텍스트                   |
| 0x7592  | `GetMessageBoxMode`                | 메시지박스 모드                 |
| 0x7593  | `SetMessageBoxMode`                | 메시지박스 모드 설정            |

## 8. 이벤트 핸들러 (IHwpObjectEvents)

`CHwpObjectEvents` 클래스가 래핑하는 이벤트:

| DISPID | 이벤트                | 설명         |
|--------|-----------------------|--------------|
| 0x01   | `Quit`                | HWP 종료     |
| 0x02   | `CreateXHwpWindow`    | 윈도우 생성  |
| 0x03   | `CloseXHwpWindow`     | 윈도우 닫기  |
| 0x04   | `NewDocument`         | 새 문서      |
| 0x05   | `DocumentBeforeClose` | 문서 닫기 전 |
| 0x06   | `DocumentBeforeOpen`  | 문서 열기 전 |
| 0x07   | `DocumentAfterOpen`   | 문서 열기 후 |
| 0x08   | `DocumentBeforeSave`  | 저장 전      |
| 0x09   | `DocumentAfterSave`   | 저장 후      |
| 0x0A   | `DocumentAfterClose`  | 문서 닫기 후 |
| 0x0B   | `DocumentChange`      | 문서 변경    |

## 9. Rust 포팅 시 참고사항

### 9.1. vtable 구현

`IHncUserActionModule`은 순수 가상 클래스 (C++ abstract class). Rust에서는 `#[repr(C)]` struct + 함수 포인터 테이블로 구현해야 함:

```
┌─────────────────────┐
│ vtable pointer ─────┼───→ ┌──────────────────┐
├─────────────────────┤     │ EnumAction       │
│ (필드들)            │     │ GetActionImage   │
└─────────────────────┘     │ UpdateUI         │
                            │ DoAction         │
                            └──────────────────┘
```

### 9.2. 호출 규약

- vtable 메서드: `extern "system"` (= `__stdcall` on x86)
- 익스포트 함수: `extern "system"` + `#[unsafe(no_mangle)]`

### 9.3. IDispatch 래핑

C++ 샘플은 MFC `COleDispatchDriver::InvokeHelper()`로 IDispatch를 호출.
Rust에서는 `windows` 크레이트의 `IDispatch::Invoke()`를 직접 사용하거나, hwp_core의 `HwpObject`에서 래핑.

### 9.4. 문자열 처리

- 액션 이름(GUID): `LPCSTR` (ANSI) — Rust에서 `CStr` / `&[u8]`
- COM 메서드 파라미터: `LPCTSTR` / `BSTR` (유니코드) — Rust에서 `HSTRING` / `BSTR`
- 액션 이름이 ANSI인 점에 주의 (UTF-8이 아님)

### 9.5. 비트맵 리소스

DLL 리소스에 비트맵을 임베딩하고 `LoadImage()`로 로드.
Rust에서는 `.rc` 파일 + `embed-resource` 크레이트 또는 Windows API 직접 사용.

### 9.6. 시리얼라이즈 경로

`ChangeSerializePath()` + `IsNewSerializePath()`로 HWP 레지스트리에 UI 상태를 저장.
이미 등록된 경우 UI를 다시 생성하지 않도록 가드.

## 10. 액션-파라미터 실행 패턴

HWP의 액션 시스템은 Action + ParameterSet 패턴:

```cpp
// CHwpObject에서 액션 생성
CDHwpAction action;
action.AttachDispatch(hwpObject.CreateAction("InsertText"));

// 파라미터 셋 생성
CDHwpParameterSet paramSet;
paramSet.AttachDispatch(action.CreateSet());

// 파라미터 설정
paramSet.SetItem("Text", COleVariant("Hello"));

// 기본값 로드 + 실행
action.GetDefault(paramSet);
action.Execute(paramSet);
```

`CDHwpAction` (DISPID):
- `0x01` get_ActID, `0x02` get_SetID
- `0x3A98` GetDefault, `0x3A99` CreateSet, `0x3A9A` Execute, `0x3A9B` PopupDialog, `0x3A9C` Run

`CDHwpParameterSet` (DISPID):
- `0x3A9D` Item(itemid), `0x3AA2` SetItem(itemid, val)
- `0x3A99` CreateItemArray, `0x3A9A` CreateItemSet
- `0x3A9C` IsEquivalent, `0x3A9F` Merge
