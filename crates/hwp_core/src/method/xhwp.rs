//! SDK `IXHwp*` — 문서·창·툴바 등 하위 객체의 method/property binding
//!
//! SDK 참고: HwpAutomation_2504.pdf, HwpObject.tlb.IDL
use crate::error::{Result, ensure};
use crate::variant::IntoVariant;

// =========================================================================
// CXHwpDocuments — HWP 문서(Document) 컬렉션
// =========================================================================

hwp_com_type!(
    /// SDK `IXHwpDocuments` 대응 — IXHwpDocument(Document) 오브젝트를 관리하는 오브젝트.
    ///
    /// SDK 참고: HwpAutomation_2504.pdf § IXHwpDocuments
    XHwpDocuments
);

impl XHwpDocuments {
    /// 현재 활성화 상태인 Document를 반환합니다.
    ///
    /// SDK: `Active_XHwpDocument(Property)` — IXHwpDocument
    pub fn active(&self) -> Result<XHwpDocument> {
        self.get("Active_XHwpDocument")
    }

    /// 열린 문서의 총 개수를 반환합니다.
    ///
    /// SDK: `Count(Property)`
    pub fn count(&self) -> Result<i32> {
        self.get("Count")
    }

    /// 인덱스로 문서를 가져옵니다 (0-based).
    ///
    /// SDK: `LPDISPATCH Item(long index)`
    pub fn item(&self, index: i32) -> Result<XHwpDocument> {
        self.get_with("Item", vec![index.into_variant()?])
    }

    /// Document 오브젝트를 추가합니다.
    ///
    /// SDK: `LPDISPATCH Add(BOOL isTab)`
    /// - `is_tab = true` → 새 탭으로 열림
    /// - `is_tab = false` → 새 창으로 열림
    pub fn add(&self, is_tab: bool) -> Result<XHwpDocument> {
        self.call_with("Add", vec![is_tab.into_variant()?])
    }

    /// Document 아이디로 지정된 Document를 찾습니다.
    ///
    /// SDK: `LPDISPATCH FindItem(long Docid)`
    pub fn find_item(&self, doc_id: i32) -> Result<XHwpDocument> {
        self.call_with("FindItem", vec![doc_id.into_variant()?])
    }
}

/// `IXHwpDocument.Modified`의 값 — 문서의 변경 상태
///
/// IDL에서 이 property는 `long`이다. `IXHwpDocument.Modified` 항목에는 값 설명이 없어,
/// 같은 뜻의 `IHwpObject.IsModified` 항목에 적힌 값을 상수로 둔다.
///
/// SDK 참고: HwpAutomation_2504.pdf § IHwpObject `IsModified(Property)`
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ModifiedState(i32);

impl ModifiedState {
    /// 변경되지 않은 깨끗한 상태 (SDK: 0)
    pub const CLEAN: Self = Self(0);
    /// 변경된 상태 (SDK: 1)
    pub const MODIFIED: Self = Self(1);
    /// 변경되었으나 자동 저장된 상태 (SDK: 2)
    pub const AUTO_SAVED: Self = Self(2);

    /// SDK의 `long` 값
    pub const fn bits(self) -> i32 {
        self.0
    }

    /// `long` 값을 그대로 감싼다. 정의되지 않은 값도 버리지 않는다.
    pub const fn from_bits_retain(bits: i32) -> Self {
        Self(bits)
    }

    /// 마지막 저장 뒤 변경됐는지. [`CLEAN`](Self::CLEAN)이 아니면 `true`다
    /// (자동 저장만 된 상태도 변경된 것으로 본다).
    pub const fn is_modified(self) -> bool {
        self.0 != Self::CLEAN.0
    }
}

/// `IXHwpDocument.EditMode`의 값 — 문서의 편집 모드
///
/// 값이 비트 조합으로 올 수 있으므로 `==` 대신 [`contains`](Self::contains)로 판단한다.
/// 배포용 문서는 16 비트로 판단한다(실측 17).
///
/// SDK 참고: HwpAutomation_2504.pdf § IHwpObject `EditMode(Property)`
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct EditMode(i32);

impl EditMode {
    /// 읽기 전용 (SDK: 0). 0이라 `contains`로는 판단할 수 없으므로
    /// [`is_read_only`](Self::is_read_only)로 판단한다.
    pub const READ_ONLY: Self = Self(0);
    /// 일반 편집 모드 (SDK: 1)
    pub const NORMAL: Self = Self(1);
    /// 양식 모드(양식 사용자 모드). 셀과 누름틀 가운데 양식 모드에서 편집 가능 속성을 가진
    /// 것만 편집할 수 있다. (SDK: 2)
    pub const FORM: Self = Self(2);
    /// 배포용 문서. SDK에 "SetEditMode로 지정 불가능"이라고 적혀 있다. (SDK: 16)
    pub const DISTRIBUTION: Self = Self(16);

    /// SDK에 넘길 `long` 값
    pub const fn bits(self) -> i32 {
        self.0
    }

    /// `long` 값을 그대로 감싼다. 정의되지 않은 값도 버리지 않는다.
    pub const fn from_bits_retain(bits: i32) -> Self {
        Self(bits)
    }

    /// `other`의 비트가 모두 켜져 있는지
    pub const fn contains(self, other: Self) -> bool {
        self.0 & other.0 == other.0
    }

    /// 읽기 전용인지. 배포용 비트(16)는 빼고 판단한다.
    ///
    /// 16 말고 다른 비트가 함께 오면 false가 된다.
    pub const fn is_read_only(self) -> bool {
        self.0 & !Self::DISTRIBUTION.0 == 0
    }
}

// =========================================================================
// CXHwpDocument — HWP 개별 문서(Document)
// =========================================================================

hwp_com_type!(
    /// SDK `IXHwpDocument` 대응 — Document 오브젝트.
    ///
    /// SDK 참고: HwpAutomation_2504.pdf § IXHwpDocument
    XHwpDocument
);

impl XHwpDocument {
    /// Document의 Path를 반환합니다 (읽기 전용).
    ///
    /// SDK: `Path(Property)`
    pub fn path(&self) -> Result<String> {
        self.get("Path")
    }

    /// Document의 전체 경로를 반환합니다 (읽기 전용).
    ///
    /// SDK: `FullName(Property)`
    pub fn full_name(&self) -> Result<String> {
        self.get("FullName")
    }

    /// Document의 편집 모드를 반환합니다.
    ///
    /// SDK: `EditMode(Property)` — [`EditMode`] 참고
    pub fn edit_mode(&self) -> Result<EditMode> {
        self.get::<i32>("EditMode").map(EditMode::from_bits_retain)
    }

    /// Document의 변경 상태를 반환합니다.
    ///
    /// SDK: `Modified(Property)` — [`ModifiedState`] 참고
    pub fn modified(&self) -> Result<ModifiedState> {
        self.get::<i32>("Modified")
            .map(ModifiedState::from_bits_retain)
    }

    /// Document의 저장된 포맷을 반환합니다 (읽기 전용).
    ///
    /// SDK: `Format(Property)`
    pub fn format(&self) -> Result<String> {
        self.get("Format")
    }

    /// Document의 고유 ID를 반환합니다 (읽기 전용).
    ///
    /// SDK: `DocumentID(Property)`
    pub fn document_id(&self) -> Result<i32> {
        self.get("DocumentID")
    }

    /// 이 Document를 활성화 상태로 만듭니다.
    ///
    /// SDK: `SetActive_XHwpDocument(Method)` — 문서를 활성화 상태로 하기
    pub fn set_active(&self) -> Result<()> {
        self.call("SetActive_XHwpDocument")
    }
}

// =========================================================================
// CXHwpWindows — HWP 창 컬렉션
// =========================================================================

hwp_com_type!(
    /// SDK `IXHwpWindows` 대응 — IXHwpWindow 오브젝트를 관리하는 오브젝트.
    ///
    /// SDK 참고: HwpAutomation_2504.pdf § IXHwpWindows
    XHwpWindows
);

impl XHwpWindows {
    /// 현재 활성 창을 반환합니다.
    pub fn active(&self) -> Result<XHwpWindow> {
        self.get("Active_XHwpWindow")
    }

    pub fn count(&self) -> Result<i32> {
        self.get("Count")
    }

    pub fn item(&self, index: i32) -> Result<XHwpWindow> {
        self.get_with("Item", vec![index.into_variant()?])
    }
}

// =========================================================================
// CXHwpWindow — HWP 개별 창
// =========================================================================

hwp_com_type!(
    /// SDK `IXHwpWindow` 대응 — 윈도우 오브젝트.
    ///
    /// SDK 참고: HwpAutomation_2504.pdf § IXHwpWindow
    XHwpWindow
);

impl XHwpWindow {
    /// 툴바 레이아웃 객체를 반환합니다.
    pub fn toolbar_layout(&self) -> Result<XHwpToolbarLayout> {
        self.get("XHwpToolbarLayout")
    }

    /// 창 표시 여부를 읽습니다.
    pub fn visible(&self) -> Result<bool> {
        self.get("Visible")
    }

    /// 창 표시 여부를 설정합니다.
    ///
    /// SDK: `XHwpWindow.Visible(Property)` — 윈도우 보이기/보이지 않기 설정/얻음
    pub fn set_visible(&self, visible: bool) -> Result<()> {
        self.put("Visible", visible)
    }

    /// 창의 좌측(X) 좌표 (픽셀).
    ///
    /// SDK: `XHwpWindow.Left(Property)` — 읽기/쓰기 가능
    pub fn left(&self) -> Result<i32> {
        self.get("Left")
    }

    /// 창의 상단(Y) 좌표 (픽셀).
    ///
    /// SDK: `XHwpWindow.Top(Property)` — 읽기/쓰기 가능
    pub fn top(&self) -> Result<i32> {
        self.get("Top")
    }

    /// 창의 가로 크기 (픽셀).
    ///
    /// SDK: `XHwpWindow.Width(Property)` — 읽기/쓰기 가능
    pub fn width(&self) -> Result<i32> {
        self.get("Width")
    }

    /// 창의 세로 크기 (픽셀).
    ///
    /// SDK: `XHwpWindow.Height(Property)` — 읽기/쓰기 가능
    pub fn height(&self) -> Result<i32> {
        self.get("Height")
    }
}

// =========================================================================
// CXHwpToolbarLayout — 툴바 레이아웃 관리
// =========================================================================

hwp_com_type!(
    /// SDK `IXHwpToolbarLayout` 대응 — 툴바나 툴바 버튼의 생성/삭제와 관련된 작업을 한다.
    ///
    /// SDK 참고: Addon-Action.pdf § 3. Toolbar 관련 개체 › 가. IXHwpToolbarLayout
    XHwpToolbarLayout
);

impl XHwpToolbarLayout {
    /// 직렬화 경로를 변경합니다. HWP가 툴바 상태를 저장/복원하는 데 사용합니다.
    pub fn change_serialize_path(&self, path: &str) -> Result<()> {
        self.call_with("ChangeSerializePath", vec![path.into_variant()?])
    }

    /// 현재 직렬화 경로가 새 것인지 확인합니다.
    /// `true`면 UI를 처음 생성해야 합니다.
    pub fn is_new_serialize_path(&self) -> Result<bool> {
        self.call("IsNewSerializePath")
    }

    /// 새 툴바를 생성합니다.
    pub fn create_toolbar(&self, name: &str, style: i32) -> Result<XHwpToolbar> {
        self.call_with(
            "CreateToolbar",
            vec![name.into_variant()?, style.into_variant()?],
        )
    }

    /// 툴바 버튼을 생성합니다.
    /// `style`: 0x01=아이콘, 0x02=텍스트, 0x03=아이콘+텍스트
    pub fn create_toolbar_button(
        &self,
        name: &str,
        action_id: &str,
        style: i32,
    ) -> Result<XHwpToolbarButton> {
        self.call_with(
            "CreateToolbarButton",
            vec![
                name.into_variant()?,
                action_id.into_variant()?,
                style.into_variant()?,
            ],
        )
    }

    pub fn get_toolbar(&self, name: &str) -> Result<XHwpToolbar> {
        self.call_with("GetToolbar", vec![name.into_variant()?])
    }

    pub fn delete_toolbar(&self, name: &str) -> Result<()> {
        let ok: bool = self.call_with("DeleteToolbar", vec![name.into_variant()?])?;
        ensure(ok, || format!("DeleteToolbar({name:?})"))
    }

    pub fn show_toolbar(&self, name: &str, show: bool) -> Result<()> {
        let ok: bool = self.call_with(
            "ShowToolbar",
            vec![name.into_variant()?, show.into_variant()?],
        )?;
        ensure(ok, || format!("ShowToolbar({name:?}, {show})"))
    }

    /// 메뉴 버튼을 생성합니다.
    /// `action_id`가 빈 문자열이면 서브메뉴 컨테이너가 됩니다.
    pub fn create_menu_button(
        &self,
        name: &str,
        action_id: &str,
        style: i32,
    ) -> Result<XHwpToolbarMenuButton> {
        self.call_with(
            "CreateMenuButton",
            vec![
                name.into_variant()?,
                action_id.into_variant()?,
                style.into_variant()?,
            ],
        )
    }

    /// ToolBox(리본) 툴바 객체를 반환합니다.
    pub fn get_toolbox_toolbar(&self) -> Result<XHncToolBoxToolbar> {
        self.call("GetToolBoxToolBar")
    }
}

// =========================================================================
// CXHwpToolbar — 개별 툴바
// =========================================================================

hwp_com_type!(
    /// SDK `IXHwpToolbar` 대응. 툴바 버튼을 추가·삭제하거나 얻어 온다.
    ///
    /// SDK 참고: Addon-Action.pdf § 3. Toolbar 관련 개체 › 나. IXHwpToolbar
    /// (개체 설명 없이 메서드 설명만 있다)
    XHwpToolbar
);

impl XHwpToolbar {
    /// 버튼을 지정 위치에 삽입합니다. `pos = -1`이면 끝에 추가합니다.
    pub fn insert_button(&self, button: XHwpToolbarButton, pos: i32) -> Result<()> {
        self.call_with(
            "InsertToolbarButton",
            vec![button.into_variant()?, pos.into_variant()?],
        )
    }

    /// 메뉴 버튼을 툴바에 삽입합니다. `pos = -1`이면 끝에 추가합니다.
    /// C++ `InsertToolbarButton(LPDISPATCH, long)`은 LPDISPATCH를 받으므로
    /// 메뉴 버튼도 삽입 가능합니다.
    pub fn insert_menu_button(&self, button: XHwpToolbarMenuButton, pos: i32) -> Result<()> {
        self.call_with(
            "InsertToolbarButton",
            vec![button.into_variant()?, pos.into_variant()?],
        )
    }

    pub fn delete_button(&self, pos: i32) -> Result<()> {
        let ok: bool = self.call_with("DeleteToolbarButton", vec![pos.into_variant()?])?;
        ensure(ok, || format!("DeleteToolbarButton({pos})"))
    }

    pub fn get_button(&self, pos: i32) -> Result<XHwpToolbarButton> {
        self.call_with("GetToolbarButton", vec![pos.into_variant()?])
    }
}

// =========================================================================
// CXHwpToolbarButton — 툴바 버튼
// =========================================================================

hwp_com_type!(
    /// SDK `IXHwpToolbarButton` 대응. 툴바 버튼이며, 툴바와 메뉴 양쪽에 쓸 수 있다.
    ///
    /// SDK 참고: Addon-Action.pdf § 3. Toolbar 관련 개체 › 다. IXHwpToolbarButton
    /// (개체 설명 없이 메서드 설명만 있다)
    XHwpToolbarButton
);

// =========================================================================
// CXHwpToolbarMenuButton — 메뉴 버튼
// =========================================================================

hwp_com_type!(
    /// SDK `IXHwpToolbarMenuButton` 대응 — IXHwpToolbarButton을 상속받으므로 IXHwpToolbarButton의
    /// 모든 함수를 사용할 수 있다. 서브 메뉴 버튼을 추가·삭제하거나 얻어 온다.
    ///
    /// SDK 참고: Addon-Action.pdf § 3. Toolbar 관련 개체 › 라. IXHwpToolbarMenuButton
    XHwpToolbarMenuButton
);

impl XHwpToolbarMenuButton {
    /// 자식 메뉴 버튼을 삽입합니다. `pos = -1`이면 끝에 추가합니다.
    pub fn insert_menu_button(&self, button: XHwpToolbarMenuButton, pos: i32) -> Result<()> {
        self.call_with(
            "InsertMenuButton",
            vec![button.into_variant()?, pos.into_variant()?],
        )
    }

    /// 툴바 버튼을 메뉴 아이템으로 삽입합니다. `pos = -1`이면 끝에 추가합니다.
    /// SDK: "한글에서는 툴바 버튼과 메뉴 버튼을 구별하지 않습니다."
    pub fn insert_toolbar_button(&self, button: XHwpToolbarButton, pos: i32) -> Result<()> {
        self.call_with(
            "InsertMenuButton",
            vec![button.into_variant()?, pos.into_variant()?],
        )
    }
}

// =========================================================================
// ToolBox(리본) 관련 타입
// =========================================================================

hwp_com_type!(
    /// ToolBox(리본)의 최상위 객체. [`XHwpToolbarLayout::get_toolbox_toolbar`]가 돌려준다.
    ///
    /// SDK 문서와 `HwpObject.tlb.IDL`에 이 인터페이스의 설명이 없다. 이름과 멤버는 C++ 샘플
    /// `HncUserActionSample/Objects/HncObjects/CXHncToolBoxToolbar.h`에서 가져왔으므로 동작은 실환경 확인이 필요하다.
    XHncToolBoxToolbar
);

impl XHncToolBoxToolbar {
    /// 탭을 삽입합니다. `index = -1`이면 끝에 추가합니다.
    pub fn insert_toolbox_tab(&self, index: i32, uid: &str, name: &str) -> Result<XHncToolBoxTab> {
        self.call_with(
            "InsertToolBoxTab",
            vec![
                index.into_variant()?,
                uid.into_variant()?,
                name.into_variant()?,
            ],
        )
    }

    /// ToolBox 아이템 버튼을 생성합니다.
    /// `btn_type`: 1=STDTB_BTN, 2=SEPARATOR, 6=COMBO, 7=MENU
    /// `style`: 0x01=아이콘, 0x02=텍스트, 0x03=아이콘+텍스트
    pub fn create_toolbox_item_button_ex(
        &self,
        name: &str,
        aid: &str,
        btn_type: i32,
        style: i32,
    ) -> Result<crate::disp_obj::DispObj> {
        self.call_with(
            "CreateToolBoxItemButtonEx",
            vec![
                name.into_variant()?,
                aid.into_variant()?,
                btn_type.into_variant()?,
                style.into_variant()?,
            ],
        )
    }
}

hwp_com_type!(
    /// ToolBox 탭.
    ///
    /// SDK 문서와 `HwpObject.tlb.IDL`에 이 인터페이스의 설명이 없다. 이름과 멤버는 C++ 샘플
    /// `HncUserActionSample/Objects/HncObjects/CXHncToolBoxTab.h`에서 가져왔으므로 동작은 실환경 확인이 필요하다.
    XHncToolBoxTab
);

impl XHncToolBoxTab {
    pub fn uid(&self) -> Result<String> {
        self.get("UID")
    }

    pub fn name(&self) -> Result<String> {
        self.get("Name")
    }

    /// 탭 내에 ToolBox를 삽입합니다.
    pub fn insert_toolbox(&self, index: i32, uid: &str, name: &str) -> Result<XHncToolBox> {
        self.call_with(
            "InsertToolBox",
            vec![
                index.into_variant()?,
                uid.into_variant()?,
                name.into_variant()?,
            ],
        )
    }
}

hwp_com_type!(
    /// 탭 안의 ToolBox.
    ///
    /// SDK 문서와 `HwpObject.tlb.IDL`에 이 인터페이스의 설명이 없다. 이름과 멤버는 C++ 샘플
    /// `HncUserActionSample/Objects/HncObjects/CXHncToolBox.h`에서 가져왔으므로 동작은 실환경 확인이 필요하다.
    XHncToolBox
);

impl XHncToolBox {
    /// 레이아웃을 가져옵니다.
    pub fn get_layout(&self, index: i32) -> Result<XHncToolBoxLayout> {
        self.call_with("GetLayout", vec![index.into_variant()?])
    }
}

hwp_com_type!(
    /// ToolBox 레이아웃.
    ///
    /// SDK 문서와 `HwpObject.tlb.IDL`에 이 인터페이스의 설명이 없다. 이름과 멤버는 C++ 샘플
    /// `HncUserActionSample/Objects/HncObjects/CXHncToolBoxLayout.h`에서 가져왔으므로 동작은 실환경 확인이 필요하다.
    XHncToolBoxLayout
);

impl XHncToolBoxLayout {
    /// 그룹을 삽입합니다.
    /// `group_type`: 0=LARGEICON, 1=LARGEICON_HORZTEXT, 2=ROW_COL
    pub fn insert_group(
        &self,
        index: i32,
        uid: &str,
        group_type: i32,
        rows: i32,
        cols: i32,
    ) -> Result<XHncToolBoxGroup> {
        self.call_with(
            "InsertGroup",
            vec![
                index.into_variant()?,
                uid.into_variant()?,
                group_type.into_variant()?,
                rows.into_variant()?,
                cols.into_variant()?,
            ],
        )
    }
}

hwp_com_type!(
    /// ToolBox 그룹.
    ///
    /// SDK 문서와 `HwpObject.tlb.IDL`에 이 인터페이스의 설명이 없다. 이름과 멤버는 C++ 샘플
    /// `HncUserActionSample/Objects/HncObjects/CXHncToolBoxGroup.h`에서 가져왔으므로 동작은 실환경 확인이 필요하다.
    XHncToolBoxGroup
);

impl XHncToolBoxGroup {
    /// 아이템을 삽입합니다. `index = -1`이면 끝에 추가합니다.
    pub fn insert_item(&self, index: i32, item: crate::disp_obj::DispObj) -> Result<i32> {
        self.call_with(
            "InsertItem",
            vec![index.into_variant()?, item.into_variant()?],
        )
    }
}
