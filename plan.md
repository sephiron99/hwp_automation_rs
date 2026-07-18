# HWP SDK Action 추가 구현 계획

## 1. 목표와 완료 범위

`hwp_sdk_doc/ActionTable_2504.pdf`를 기준으로 Automation에서 사용할 수 있는 Action을 빠짐없이
`hwp_core`에 노출한다. `ParameterSetTable_2504.pdf`, `HwpAutomation_2504.pdf`,
`HwpObject.tlb.IDL`은 실행 방식과 COM 타입을 확인하는 교차 검증 자료로 사용한다.

완료 목표는 두 단계로 나눈다.

1. **Action 완전성**
   - 실행 가능한 모든 Action ID에 이름 있는 Rust 진입점과 메타데이터가 있다.
   - 파라미터 Action도 범용 `IDHwpAction`/`IDHwpParameterSet` API로 실행할 수 있다.
   - 실행할 수 없는 Action은 SDK 근거와 함께 제외 목록에 들어간다.
2. **타입 안전 편의 API**
   - 자주 사용하는 ParameterSet부터 전용 Rust 타입과 편의 메서드를 제공한다.
   - 최종적으로 Action이 직접 또는 하위 Set으로 참조하는 공개 ParameterSet을 모두 바인딩한다.

Dummy Action, HwpCtrl 전용 Action, 외부에 ParameterSet이 노출되지 않아 Automation에서 실행할 수
없는 Action은 “구현 누락”이 아니라 근거가 있는 “미지원”으로 관리한다.

## 2. 재검토한 현재 상태

- 현재 `crates/hwp_core/src/actions/`에는 14개 도메인 파일이 있고,
  `HAction::run("...")` 기준 고유 Action ID는 약 530개다.
- 공개 Action 메서드는 대부분 `self.h_action()?.run("ActionId")`만 호출한다.
- `actions/mod.rs`와 `hwp_core.md`에는 139개 이동 Action을 가진 `move_` 모듈이 있다고 적혀 있지만,
  실제 `actions/move_.rs`와 `pub mod move_;`는 없다.
- `HParameterSet`의 전용 접근자는 `HInsertText`, `HCharShape`, `HParaShape` 세 종류뿐이다.
- `ParameterSetTable_2504.pdf`에는 141개의 최상위 정의가 있지만, 모든 정의가 Action의 직접
  ParameterSet은 아니다. 일부는 다른 Set의 `PIT_SET`/`PIT_ARRAY` 하위 타입이다.
- PDF에서 Action ID 첫 번째 열을 기계 추출하면 993개 조각이 나온다. 긴 ID의 줄바꿈과
  `AutoSpellSelect1 ~ 16` 같은 범위 표기가 포함되므로, 정규화 전에는 총 Action 수로 간주할 수 없다.
- 현재 코드의 일부 Action ID는 PDF 원문과 철자가 다르거나 PDF에서 Dummy로 표시된다. 기존 API도
  신규 인벤토리에 넣어 호환성과 실동작을 함께 확인해야 한다.

## 3. SDK 실행 모델

Action 래퍼를 늘리기 전에 아래 네 계층을 구분해야 한다.

| SDK 계층 | 역할 | 현재 상태 | 구현 방침 |
|---|---|---|---|
| `IHwpObject.Run(action)` | Action ID를 직접 실행 | `HwpObject::run` 구현됨 | 문자열 기반 최저 수준 API로 유지 |
| `HAction` 프로퍼티 | 이름을 받아 `GetDefault`, `Execute`, `PopupDialog`, `Run` 수행 | 기본 메서드 구현됨 | 기존 API 호환 유지, 문서 의미 수정 |
| `IHwpObject.CreateAction(action)` → `IDHwpAction` | 특정 Action 핸들과 그 Action의 Set 생성 | 미구현 | ParameterSet Action의 범용 실행 기반으로 우선 구현 |
| `IDHwpParameterSet`/`IDHwpParameterArray` | 이름 기반 Set/Array 항목 조작 | 미구현 | 중첩 ParameterSet까지 다룰 수 있게 우선 구현 |

SDK에서 `HAction.Run`은 단순한 “무파라미터 Action 실행”이 아니라 `GetDefault`, `PopupDialog`,
`Execute`를 한 번에 수행하는 호출로 설명된다. 따라서 다음 원칙을 적용한다.

- `ParameterSet ID`가 `-`인 Action은 직접 실행 편의 메서드를 제공한다.
- 공개 ParameterSet을 가진 Action은 대화상자/기본 실행과 프로그래밍 방식 실행을 구분한다.
- 파라미터 Action을 인자 없는 `run`으로 감싼 기존 메서드는 즉시 제거하지 않고, 실제 의미를
  “기본 UI 실행”으로 문서화한다.
- 자동화에 적합한 실행은 `CreateAction → CreateSet → GetDefault → SetItem → Execute` 경로를
  기본으로 한다.

## 4. 인벤토리와 지원 여부 판정

### 4.1 인벤토리 필드

PDF에서 다음 필드를 추출해 저장소에 정규화된 인벤토리 파일로 체크인한다.

- 원본 Action ID와 Rust 메서드명
- ParameterSet ID 및 `-`, `+`, `*` 표기
- 설명, 비고, 명시된 HWP 버전
- PDF 스타일에 따른 Dummy/HwpCtrl 전용 여부
- 구현 모듈, 실행 경로, 구현 상태, 제외 사유

PDF는 gitignored이므로 일반 빌드와 검증은 PDF 파일에 의존하지 않는다. PDF 추출기는 인벤토리 갱신용
도구로만 사용한다. 정규화 결과는 한 행이 한 Action인 `data/actions.csv`, PDF 추출의 예외와 수동
판정은 근거를 주석으로 남길 수 있는 `data/action_overrides.toml`에 저장한다.
설명·비고 필드에 쉼표가 포함된 한국어 문장이 들어가므로 CSV는 RFC 4180 큰따옴표 quoting을
따르거나, 파싱 단순화를 위해 탭 구분(TSV)을 사용한다. 추출 도구와 검증 도구가 같은 규칙을
공유해야 한다.

### 4.2 정규화 규칙

1. 페이지가 바뀌거나 열 너비 때문에 나뉜 Action ID를 결합한다.
2. `AutoSpellSelect1 ~ 16`과 같은 범위는 실제 ID 16개로 확장한다.
3. ID 내부의 의미 있는 공백과 SDK 오탈자는 COM 문자열에서 그대로 보존한다.
4. Rust 메서드명만 일관된 `snake_case`로 변환한다.
5. `pdftotext`가 잃어버리는 글자색/기울임은 PDF XML의 폰트 스타일로 추출하고 표본을 수동 확인한다.
6. 같은 ID가 여러 행에 나타나면 중복인지 버전별 정의인지 비고까지 비교한다.

### 4.3 두 축의 분류

지원 여부와 실행 방식을 하나의 상태로 섞지 않고 별도 축으로 관리한다.

지원 여부:

| 상태 | SDK 조건 | 처리 |
|---|---|---|
| `automation` | 일반 Action | 구현 대상 |
| `dummy` | 빨간색 밑줄 | 신규 구현 제외, 기존 래퍼는 실동작/호환성 조사 |
| `hwpctrl-only` | 기울임 밑줄 | `HwpObject` 구현 대상에서 제외 |
| `unexposed` | `ParameterSet ID`가 `+` | 범용 실행 가능 여부를 Windows에서 확인하기 전까지 보류 |
| `dependent` | `ParameterSet ID`가 `*` | 필요한 상위 Action/Set 생성 경로가 확인된 경우에만 구현 |

실행 방식:

| 상태 | 조건 | 공개 API |
|---|---|---|
| `direct` | ParameterSet 없음 | 인자 없는 Action 편의 메서드 |
| `interactive` | 기본 대화상자/기본값 실행이 유효 | 기존 인자 없는 메서드 또는 명시적인 `_dialog` 메서드 |
| `generic-set` | 공개 ParameterSet 사용 | `ActionHandle` + `ParameterSet` 범용 API |
| `typed-set` | 타입 안전 래퍼 구현 완료 | 전용 파라미터 타입/빌더와 편의 메서드 |

각 정규화 Action은 지원 여부 한 개와 실행 방식 한 개 이상을 반드시 가져야 한다.

Action별 Rust 진입점은 실행 방식에 따라 다음 형태로 통일한다.

- `direct`: `foo(&self) -> Result<()>`
- `interactive`: `foo_dialog(&self) -> Result<()>`; 이미 공개된 `foo()`는 호환 별칭으로 유지
- `generic-set`: `foo_action(&self) -> Result<ActionHandle>`
- `typed-set`: `foo_with(&self, params: &FooParams) -> Result<()>` 또는 동등한 전용 API

Rust는 메서드 오버로딩을 지원하지 않으므로 접미사를 생략하지 않는다. 구체적인 이름은 기존 공개
API와의 충돌을 확인한 뒤 인벤토리에 고정한다.

## 5. 구현 단계

### 단계 A. 재현 가능한 SDK 인벤토리 작성

1. Action Table 추출 도구와 수동 보정 파일을 추가한다.
2. 현재 Rust 소스의 `run`, `get_default`, `execute` 문자열을 추출한다.
3. 다음 차이를 생성한다.
   - `implemented`
   - `missing`
   - `name-mismatch`
   - `duplicate`
   - `dummy-existing`
   - `excluded`
4. 기존 530개 ID를 포함하여 모든 항목에 근거 페이지와 상태를 부여한다.
5. `hwp_core.md`의 부정확한 메서드 수는 실제 상태로 바로잡는다. 존재하지 않는 `move_` 모듈
   기술은 삭제하지 않고 "(미구현, 단계 D에서 추가 예정)"으로 표기해 단계 D와의 이중 수정을
   피한다.

완료 조건:

- 정규화된 전체 Action 수가 확정되고 원본 PDF 표본과 일치한다.
- 모든 현재 Action 메서드가 인벤토리의 정확히 한 항목과 대응한다.
- 소스와 인벤토리의 누락/중복을 검출하는 검증 명령이 있다.

### 단계 B. 범용 Action/ParameterSet COM 계층 구현

타입별 ParameterSet을 대량 구현하기 전에 IDL에 이미 정의된 범용 인터페이스를 노출한다.

1. `ActionHandle` (`IDHwpAction`)을 추가한다.
   - `act_id`, `set_id`
   - `create_set`
   - `get_default`
   - `execute`
   - `popup_dialog`
   - `run`
2. `ParameterSet` (`IDHwpParameterSet`)을 추가한다.
   - `count`, `is_set`, `set_id`, `clone_set`
   - `item`, `item_exists`, `set_item`, `remove_item`, `remove_all`
   - `create_item_set`, `create_item_array`, `merge`, `get_intersection`, `is_equivalent`, `copy`
3. `ParameterArray` (`IDHwpParameterArray`)를 추가한다.
   - `count` 읽기/쓰기
   - `item`, `set_item`, `clone_array`, `copy`
4. `HwpObject`에 다음 메서드를 추가한다.
   - `create_action`
   - `create_set`
   - `release_action`
   - `is_action_enabled` (COM 메서드명은 `IsActionEnable`이다. 끝에 `d`가 없으므로 호출
     문자열에 주의한다)
5. `PIT_I1`, `PIT_I2`, `PIT_I4`, `PIT_UI1`, `PIT_UI2`, `PIT_UI4`, `PIT_UI64`, `PIT_BSTR`,
   `PIT_SET`, `PIT_ARRAY`, `PIT_BINDATA`를 다룰 수 있도록 `Variant` 변환 범위를 점검하고 보강한다.
6. 기존 `HAction`과 새 `ActionHandle`의 이름과 역할을 문서에서 명확히 구분한다.

완료 조건:

- 전용 타입이 없는 공개 ParameterSet Action도 문자열 필드 기반으로 실행할 수 있다.
- 중첩 Set과 Array를 생성하고 값을 읽고 쓸 수 있다.
- IDL의 반환 타입(`bool`, `long`, `IDispatch`)과 Rust 반환 타입이 일치한다.

### 단계 C. Action 선언 구조와 기존 API 정리

1. 반복되는 Action 메서드를 선언형 테이블 또는 내부 매크로로 관리한다.
2. 한 선언에서 다음을 연결한다.
   - Rust 메서드명
   - 정확한 COM Action ID
   - ParameterSet ID
   - 실행 방식
   - SDK 설명과 명시된 최소 버전
3. 기존 공개 메서드명과 동작은 가능한 한 유지한다.
4. 기존 ID가 잘못됐더라도 대체 ID의 실동작이 확인되기 전에는 조용히 바꾸지 않는다.
5. 이름을 바로잡아야 하면 기존 메서드는 deprecated 별칭으로 남기고 마이그레이션 문서를 제공한다.
6. 신규 버전 여부는 PDF/IDL에 근거가 있을 때만 정적으로 제한한다. 근거가 불명확하면
   `HwpVer` 추측 대신 `is_action_enabled`와 실제 실행 결과를 사용한다.

완료 조건:

- 신규 단순 Action은 한 개의 선언으로 메서드와 메타데이터에 함께 추가된다.
- 기존 사용자의 소스 호환성을 유지한다.
- Action ID 중복과 Rust 메서드명 충돌이 검출된다.
- 선언형 매크로/테이블로 이관한 뒤에도 현재의 메서드별 SDK 인용 doc comment가 보존된다
  (매크로 인자로 doc 문자열을 전달). 매크로 생성 메서드의 rust-analyzer 탐색성 저하는
  이관 전에 감수 여부를 결정한다.

### 단계 D. 누락 Action 메서드 구현

1. 가장 먼저 `actions/move_.rs`를 추가해 이동/선택 Action을 복구하고 `actions/mod.rs`에서 공개한다.
2. 기존 모듈을 다음 순서로 보강한다.
   - 편집, 삭제, 클립보드, 찾기
   - 글자 모양, 문단 모양, 스타일, 변환
   - 표와 셀
   - 페이지, 구역, 바탕쪽, 머리말/꼬리말, 주석
   - 그리기 개체, 그림, 수식, 양식
   - 파일, 보기, 변경 추적, 매크로
3. 큰 독립 영역은 `convert.rs`, `field.rs`, `form.rs`, `equation.rs` 같은 새 모듈로 분리한다.
4. `direct` Action은 인자 없는 편의 메서드를 추가한다.
5. `interactive` Action은 대화상자를 띄우거나 기본값으로 실행한다는 사실이 메서드명/문서에 드러나게 한다.
6. 공개 ParameterSet Action은 최소한 이름 있는 `foo_action()` 생성 메서드를 제공한다. 전용 편의 API
   구현 여부와 무관하게 이 진입점과 범용 Set 실행 예제가 있어야 Action 자체를 완료 처리한다.

완료 조건:

- 모든 `automation` Action에 이름 있는 실행 메서드 또는 `ActionHandle` 생성 메서드가 있다.
- 모든 미구현 항목은 기술적 보류 사유를 가진다.
- 근거 없이 모든 Action을 `HAction::run`으로 감싼 메서드는 추가하지 않는다.

### 단계 E. 타입 안전 ParameterSet 구현

Action Table이 직접 참조하는 ParameterSet부터 구현하고 `PIT_SET`/`PIT_ARRAY`로 연결된 하위 타입을
재귀적으로 추가한다. PDF의 141개 정의를 번호순으로 무조건 구현하지 않는다.

1. 공통 생성 구조
   - Set 타입 선언과 `HSet`/범용 `ParameterSet` 접근 코드를 매크로로 정리한다.
   - 필드별 읽기/쓰기 가능 여부, 단위, 기본값 적용 순서를 기록한다.
   - 고정된 값 집합은 `enum`/bitflags/newtype을 사용하고 알 수 없는 미래 값도 보존할 수 있게 한다.
2. 1차: 핵심 편집
   - `FindReplace`, `GotoE`, `CharShape`, `ParaShape`, `InsertText`, `InsertFile`, `FileOpen`, `FileSaveAs`
3. 2차: 문서 구조
   - `SecDef`, `PageDef`, `PageBorderFill`, `HeaderFooter`, `FootnoteShape`, `AutoNum`, `PageNumPos`
4. 3차: 표와 개체
   - `TableCreation`, `Table`, `TableSplitCell`, `CellBorderFill`, `ShapeObject`와 관련 하위 Draw Set
5. 4차: 나머지 Action 참조 Set
   - 필드, 변환, 스타일, 수식, 양식, 인쇄, 보안, 그림 효과 등
6. 전용 편의 메서드는 `GetDefault → 값 설정 → Execute` 흐름을 캡슐화하되, 고급 사용자가 범용
   `ParameterSet`에도 접근할 수 있게 한다.
7. 기존 `HInsertText`, `HCharShape`, `HParaShape` 공개 API는 유지하면서 공통 구조로 이관한다.

완료 조건:

- Action이 직접 참조하는 공개 ParameterSet과 필요한 하위 Set/Array가 모두 바인딩된다.
- PDF의 Item ID, PIT 타입, SubType과 Rust 스키마가 일치한다.
- 범용 API와 타입 안전 API로 같은 Action을 실행할 수 있다.

### 단계 F. 특수/실행 불가 Action 정리

1. `+` Action은 `CreateAction`/`create_set`으로도 Set을 만들 수 없는지 Windows에서 확인한다.
2. `*` Action은 필요한 상위 Action이나 읽기 전용 Set 획득 경로를 기록한다.
3. Dummy Action은 신규 공개 API에서 제외한다.
4. HwpCtrl 전용 Action은 향후 별도 HwpCtrl 크레이트/모듈의 후보로만 기록한다.
5. 기존 Dummy/잘못된 ID 래퍼는 다음 순서로 처리한다.
   - `is_action_enabled` 확인
   - 격리된 문서에서 실동작 확인
   - 대체 ID 확인
   - 호환 별칭 및 deprecated 처리
6. 삭제는 다음 주요 버전에서만 검토하고, 이번 작업에서는 소스 호환성을 우선한다.
7. 제외 항목은 `UNSUPPORTED_ACTIONS.md`에 원본 ID, PDF 페이지, 표기, 사유를 남긴다.

완료 조건:

- 모든 정규화 Action이 구현 또는 근거가 있는 제외 상태 중 하나다.
- `+`, `*`, Dummy, HwpCtrl 전용 항목이 일반 지원 Action과 섞이지 않는다.

### 단계 G. 검증과 문서화

1. 정적 검증
   - 인벤토리와 Action 선언의 양방향 일치
   - ID/메서드 중복 검사
   - ParameterSet Item ID/PIT 타입/SubType 스키마 검사
   - `cargo fmt --all -- --check`
   - `cargo check -p hwp_core --target i686-pc-windows-msvc`
   - 워크스페이스 전체 `cargo check --target i686-pc-windows-msvc`
     (hwp_com, hwp_addon, hwp_dabbrev, 예제 크레이트의 소스 호환성 검증 겸용)
   - 변경 범위가 정리된 뒤 `cargo clippy -p hwp_core --target i686-pc-windows-msvc`
2. Windows/HWP 실환경 검증
   - HWP 2018/2022/2024에서 `is_action_enabled` 결과 기록
   - 새 임시 문서에서 안전한 `direct` Action 표본 실행
   - 대표 ParameterSet별 `CreateSet`, `GetDefault`, `SetItem`, `Execute` 검증
   - 중첩 Set/Array와 타입 안전 래퍼 검증
   - 대화상자 Action은 자동 테스트에서 분리한 수동 체크리스트로 검증
   - 종료, 파일 덮어쓰기, 삭제, 보안 설정은 별도 임시 파일과 명시적 허용 하에서만 검증
3. 문서
   - `hwp_core.md`의 실제 모듈과 메서드 수
   - 범용 `ActionHandle`/`ParameterSet` 예제
   - 타입 안전 ParameterSet 예제
   - HWP 버전별 확인 결과
   - 미지원 Action과 사유

## 6. 커밋 분할

1. 정규화 인벤토리, PDF 추출 도구, 차이 검증
2. `ActionHandle`, `ParameterSet`, `ParameterArray`, `Variant` 보강
3. Action 선언 매크로/테이블과 기존 API 이관
4. `move_` 및 편집 계열 누락 Action
5. 나머지 누락 Action을 도메인별로 분할
6. 타입 안전 ParameterSet 1차
7. 타입 안전 ParameterSet 2~4차
8. Dummy/특수 Action 및 호환 별칭 정리
9. Windows 검증 결과와 최종 문서

각 커밋은 최소한 포맷 검사, 워크스페이스 전체 크로스 체크(`cargo check --target
i686-pc-windows-msvc`), 인벤토리 일치 검사를 통과해야 한다.

## 7. 최종 완료 정의

- Action Table의 모든 정규화 ID가 구현 또는 근거가 있는 제외 상태다.
- 모든 Automation Action에 이름 있는 편의 메서드나 `ActionHandle` 생성 메서드가 있다.
- Action이 참조하는 공개 ParameterSet과 필요한 하위 Set/Array가 Rust API로 노출된다.
- Dummy/HwpCtrl 전용/외부 미노출 Action이 일반 지원 API에 잘못 포함되지 않는다.
- 기존 공개 API 호환성을 유지하고 포맷, 크로스 빌드, 인벤토리 검사가 통과한다.
- 지원 대상 HWP 버전의 대표 Action에 대해 Windows 실동작 결과가 기록된다.
