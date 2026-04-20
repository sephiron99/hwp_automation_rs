use crate::error::Result;
use crate::variant::IntoVariant;

/// C++ COleDispatchDriver 래퍼 클래스에 대응하는 Rust newtype을 생성합니다.
///
/// 생성되는 구현:
/// - `Deref<Target = DispObj>` — 범용 `get`/`call`/`call_with` 접근
/// - `FromVariant` — COM 메서드 반환값에서 자동 변환
/// - `IntoVariant` — COM 메서드 인자로 전달 가능
macro_rules! hwp_com_type {
    ($(#[$meta:meta])* $name:ident) => {
        $(#[$meta])*
        pub struct $name(crate::disp_obj::DispObj);

        impl std::ops::Deref for $name {
            type Target = crate::disp_obj::DispObj;
            fn deref(&self) -> &Self::Target {
                &self.0
            }
        }

        impl crate::variant::FromVariant for $name {
            fn from_variant(
                v: &windows::Win32::System::Variant::VARIANT,
            ) -> crate::error::Result<Self> {
                Ok($name(crate::variant::FromVariant::from_variant(v)?))
            }
        }

        impl crate::variant::IntoVariant for $name {
            fn into_variant(
                self,
            ) -> crate::error::Result<windows::Win32::System::Variant::VARIANT> {
                self.0.into_variant()
            }
        }
    };
}

// =========================================================================
// CXHwpDocuments — HWP 문서(도큐먼트) 컬렉션
// =========================================================================

hwp_com_type!(
    /// C++ `CXHwpDocuments` 대응. 열린 문서(탭) 목록을 관리합니다.
    ///
    /// SDK: `IXHwpDocuments` — Document를 관리하는 Collection 개체
    XHwpDocuments
);

impl XHwpDocuments {
    /// 현재 활성화 상태인 도큐먼트를 반환합니다.
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

    /// 도큐먼트 오브젝트를 추가합니다.
    ///
    /// SDK: `LPDISPATCH Add(BOOL isTab)`
    /// - `is_tab = true` → 새 탭으로 열림
    /// - `is_tab = false` → 새 창으로 열림
    pub fn add(&self, is_tab: bool) -> Result<XHwpDocument> {
        self.call_with("Add", vec![is_tab.into_variant()?])
    }

    /// 도큐먼트 아이디로 지정된 도큐먼트를 찾습니다.
    ///
    /// SDK: `LPDISPATCH FindItem(long Docid)`
    pub fn find_item(&self, doc_id: i32) -> Result<XHwpDocument> {
        self.call_with("FindItem", vec![doc_id.into_variant()?])
    }
}

// =========================================================================
// CXHwpDocument — HWP 개별 문서(도큐먼트)
// =========================================================================

hwp_com_type!(
    /// C++ `CXHwpDocument` 대응. 개별 문서 객체입니다.
    ///
    /// SDK: `IXHwpDocument` — Document 개체
    XHwpDocument
);

impl XHwpDocument {
    /// 도큐먼트의 Path를 반환합니다 (읽기 전용).
    ///
    /// SDK: `Path(Property)`
    pub fn path(&self) -> Result<String> {
        self.get("Path")
    }

    /// 도큐먼트의 전체 경로를 반환합니다 (읽기 전용).
    ///
    /// SDK: `FullName(Property)`
    pub fn full_name(&self) -> Result<String> {
        self.get("FullName")
    }

    /// 도큐먼트의 에디트 모드를 반환합니다.
    ///
    /// SDK: `EditMode(Property)` — 0: 읽기, 1: 일반 편집, 2: 양식 모드, 16: 배포용
    pub fn edit_mode(&self) -> Result<i32> {
        self.get("EditMode")
    }

    /// 도큐먼트의 변경 여부를 반환합니다.
    ///
    /// SDK: `Modified(Property)`
    pub fn modified(&self) -> Result<bool> {
        self.get("Modified")
    }

    /// 도큐먼트의 저장된 포맷을 반환합니다 (읽기 전용).
    ///
    /// SDK: `Format(Property)`
    pub fn format(&self) -> Result<String> {
        self.get("Format")
    }

    /// 도큐먼트의 고유 ID를 반환합니다 (읽기 전용).
    ///
    /// SDK: `DocumentID(Property)`
    pub fn document_id(&self) -> Result<i32> {
        self.get("DocumentID")
    }

    /// 이 도큐먼트를 활성화 상태로 만듭니다.
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
    /// C++ `CXHwpWindows` 대응. HWP 창 목록을 관리합니다.
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
    /// C++ `CXHwpWindow` 대응. 개별 HWP 창을 제어합니다.
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
    /// C++ `CXHwpToolbarLayout` 대응. 툴바 생성/삭제/조회를 담당합니다.
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

    pub fn delete_toolbar(&self, name: &str) -> Result<bool> {
        self.call_with("DeleteToolbar", vec![name.into_variant()?])
    }

    pub fn show_toolbar(&self, name: &str, show: bool) -> Result<bool> {
        self.call_with(
            "ShowToolbar",
            vec![name.into_variant()?, show.into_variant()?],
        )
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
    /// C++ `CXHwpToolbar` 대응. 툴바에 버튼을 삽입/삭제합니다.
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

    pub fn delete_button(&self, pos: i32) -> Result<bool> {
        self.call_with("DeleteToolbarButton", vec![pos.into_variant()?])
    }

    pub fn get_button(&self, pos: i32) -> Result<XHwpToolbarButton> {
        self.call_with("GetToolbarButton", vec![pos.into_variant()?])
    }
}

// =========================================================================
// CXHwpToolbarButton — 툴바 버튼
// =========================================================================

hwp_com_type!(
    /// C++ `CXHwpToolbarButton` 대응. 툴바 버튼 객체입니다.
    XHwpToolbarButton
);

// =========================================================================
// CXHwpToolbarMenuButton — 메뉴 버튼
// =========================================================================

hwp_com_type!(
    /// C++ `CXHwpToolbarMenuButton` 대응. 메뉴 버튼/서브메뉴 컨테이너입니다.
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
    /// C++ `CXHncToolBoxToolbar` 대응. ToolBox(리본) 최상위 객체입니다.
    XHncToolBoxToolbar
);

impl XHncToolBoxToolbar {
    /// 탭을 삽입합니다. `index = -1`이면 끝에 추가합니다.
    pub fn insert_toolbox_tab(
        &self,
        index: i32,
        uid: &str,
        name: &str,
    ) -> Result<XHncToolBoxTab> {
        self.call_with(
            "InsertToolBoxTab",
            vec![index.into_variant()?, uid.into_variant()?, name.into_variant()?],
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
    /// C++ `CXHncToolBoxTab` 대응. ToolBox 탭입니다.
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
    pub fn insert_toolbox(
        &self,
        index: i32,
        uid: &str,
        name: &str,
    ) -> Result<XHncToolBox> {
        self.call_with(
            "InsertToolBox",
            vec![index.into_variant()?, uid.into_variant()?, name.into_variant()?],
        )
    }
}

hwp_com_type!(
    /// C++ `CXHncToolBox` 대응. 탭 내의 ToolBox 컨테이너입니다.
    XHncToolBox
);

impl XHncToolBox {
    /// 레이아웃을 가져옵니다.
    pub fn get_layout(&self, index: i32) -> Result<XHncToolBoxLayout> {
        self.call_with("GetLayout", vec![index.into_variant()?])
    }
}

hwp_com_type!(
    /// C++ `CXHncToolBoxLayout` 대응. ToolBox 레이아웃입니다.
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
    /// C++ `CXHncToolBoxGroup` 대응. ToolBox 그룹입니다.
    XHncToolBoxGroup
);

impl XHncToolBoxGroup {
    /// 아이템을 삽입합니다. `index = -1`이면 끝에 추가합니다.
    pub fn insert_item(
        &self,
        index: i32,
        item: crate::disp_obj::DispObj,
    ) -> Result<i32> {
        self.call_with(
            "InsertItem",
            vec![index.into_variant()?, item.into_variant()?],
        )
    }
}
