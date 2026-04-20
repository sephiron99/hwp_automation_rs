use crate::shortcut::ShortcutKey;
use crate::toolbar::ToolbarBitmap;
use hwp_core::error::Result;
use hwp_core::hwp_obj::HwpObject;

/// HWP 예약 액션 GUID — 최초 등록 시 한 번만 호출
pub const UUIDSTR_ON_INITIAL_LOAD: &str = "{B91A2981-A001-44a9-933F-5BF70A747967}";
/// HWP 예약 액션 GUID — 창이 열릴 때마다 호출
pub const UUIDSTR_ON_LOAD: &str = "{3E4DC866-051C-4989-820E-DADC1E6264B9}";

/// 툴바/리본 중 어느 쪽을 표시할지 선택 (최소 하나 필수)
pub enum ToolbarTarget {
    /// 기존 툴바만
    Toolbar,
    /// 새 탭을 생성합니다. 탭 UID는 `serialize_path`를 사용합니다.
    /// 인자: 탭 라벨 (e.g., "추가 기능")
    ///
    /// HWP SDK는 기존 내장 탭(편집, 서식 등) 접근을 지원하지 않으므로
    /// 새 탭을 생성하는 것이 유일한 방법입니다.
    Ribbon(&'static str),
    /// 툴바 + 리본 모두. 인자: 리본 탭 라벨 (e.g., "추가 기능")
    Both(&'static str),
}

/// 플러그인의 툴바/리본 설정
pub struct ToolbarConfig {
    /// 툴바 이름 (e.g., "dabbrev")
    pub name: &'static str,
    /// HWP 직렬화 경로 (e.g., "HwpDabbrev") — 새 리본 탭 UID로도 사용됩니다.
    pub serialize_path: &'static str,
    /// 960×16, 32bit RGBA BMP 데이터 (`include_bytes!`)
    pub bitmap_data: &'static [u8],
    /// 표시 대상 (툴바 / 리본 / 둘 다)
    pub target: ToolbarTarget,
    /// 탭 내 ToolBox 삽입 위치 (`InsertToolBox`의 `index` 인자).
    /// `-1`이면 끝에 추가, `0`이면 맨 앞, `n`이면 n번째 위치입니다.
    pub ribbon_toolbox_index: i32,
}

/// 사용자 액션의 메타데이터
#[derive(Clone)]
pub struct ActionMeta {
    /// 액션 식별자 (e.g., "DabbrevExpand")
    pub name: &'static str,
    /// 버튼/메뉴 표시 텍스트 (e.g., "dabbrev")
    pub label: &'static str,
    /// 비트맵 스트립 내 아이콘 인덱스
    pub image_index: i32,
    /// 단축키 (없으면 `None`)
    pub shortcut: Option<ShortcutKey>,
}

/// 한글 플러그인(Add-in) 개발자가 구현하는 트레잇
pub trait HwpUserAction: Sync + Send {
    // ── 라이프사이클 훅 (선택, 기본 구현 제공) ──

    /// 최초 등록 시 한 번 호출됩니다.
    fn on_initial_load(&self, _hwp: &HwpObject) -> Result<bool> {
        Ok(true)
    }

    /// HWP 창이 열릴 때마다 호출됩니다.
    /// 기본 구현은 `setup_toolbar`를 호출합니다.
    fn on_load(&self, hwp: &HwpObject) -> Result<bool> {
        self.setup_toolbar(hwp)
    }

    // ── 사용자 구현 필수 ──

    /// 사용자 액션 목록을 반환합니다. 예약 GUID는 자동 포함되므로 넣지 않습니다.
    fn actions(&self) -> &'static [ActionMeta];

    /// 사용자 액션이 실행될 때 호출됩니다.
    fn do_action(&self, action_name: &str, hwp: &HwpObject) -> Result<bool>;

    // ── 선택 구현 ──

    /// 툴바/리본 자동 설정에 사용할 설정을 반환합니다.
    /// `None`이면 자동 설정을 하지 않습니다.
    fn toolbar_config(&self) -> Option<&ToolbarConfig> {
        None
    }

    fn update_ui(&self, _action_name: &str, _hwp: &HwpObject) -> u32 {
        0
    }

    /// 비트맵 핸들을 반환합니다.
    /// `toolbar_config`를 구현하면 자동으로 동작합니다.
    fn toolbar_bitmap(&self, _state: u32) -> Option<isize> {
        let config = self.toolbar_config()?;
        if config.bitmap_data.is_empty() {
            return None;
        }
        static TOOLBAR: ToolbarBitmap = ToolbarBitmap::new();
        TOOLBAR.load_from_bytes(config.bitmap_data);
        TOOLBAR.image(0).map(|(h, _)| h)
    }

    /// 툴바와 리본(ToolBox)을 자동 설정합니다.
    /// `on_load`에서 호출됩니다.
    fn setup_toolbar(&self, hwp: &HwpObject) -> Result<bool> {
        let Some(config) = self.toolbar_config() else {
            return Ok(true);
        };

        let layout = hwp.windows()?.active()?.toolbar_layout()?;
        layout.change_serialize_path(config.serialize_path)?;
        if !layout.is_new_serialize_path()? {
            return Ok(true);
        }

        let has_bitmap = !config.bitmap_data.is_empty();
        let btn_style = if has_bitmap { 3 } else { 2 };

        // ── 툴바 ──
        if matches!(
            config.target,
            ToolbarTarget::Toolbar | ToolbarTarget::Both(_)
        ) {
            let toolbar = layout.create_toolbar(config.name, 0)?;
            for action in self.actions() {
                let button = layout.create_toolbar_button(action.label, action.name, btn_style)?;
                toolbar.insert_button(button, -1)?;
            }
        }

        // ── 리본(ToolBox) ──
        if let ToolbarTarget::Ribbon(label) | ToolbarTarget::Both(label) = &config.target {
            let toolbox_toolbar = layout.get_toolbox_toolbar()?;
            let tab = toolbox_toolbar.insert_toolbox_tab(-1, config.serialize_path, label)?;

            let box_uid = config.name;
            let tbox = tab.insert_toolbox(config.ribbon_toolbox_index, box_uid, config.name)?;
            let tbox_layout = tbox.get_layout(0)?;
            let cols = self.actions().len() as i32;
            let group = tbox_layout.insert_group(-1, config.name, 2, 1, cols)?;
            for action in self.actions() {
                let item = toolbox_toolbar.create_toolbox_item_button_ex(
                    action.label,
                    action.name,
                    1,
                    btn_style,
                )?;
                group.insert_item(-1, item)?;
            }
        }

        Ok(true)
    }

    // ── 프레임워크 디스패치 (기본 구현, 오버라이드 불필요) ──

    /// 예약 GUID + 사용자 액션을 순서대로 열거합니다.
    fn enum_action(&self, index: i32) -> Option<&'static str> {
        match index as usize {
            0 => Some(UUIDSTR_ON_INITIAL_LOAD),
            1 => Some(UUIDSTR_ON_LOAD),
            i => self.actions().get(i - 2).map(|a| a.name),
        }
    }

    /// 액션 이름으로 툴바 비트맵 핸들과 해당 액션의 이미지 인덱스를 반환합니다.
    fn get_action_image(&self, action_name: &str, state: u32) -> Option<(isize, i32)> {
        let entry = self.actions().iter().find(|e| e.name == action_name)?;
        let hbitmap = self.toolbar_bitmap(state)?;
        Some((hbitmap, entry.image_index))
    }

    /// 예약 GUID → 라이프사이클 훅, 그 외 → `do_action`으로 라우팅합니다.
    ///
    /// 사용자 액션 진입 직전에 [`crate::ime::commit_composition`]을 호출해
    /// IME 조합 중 단축키 실행 시 후속 `Move*` 액션이 캐럿을 움직이지 못하는
    /// 문제를 원천 차단합니다. (툴바 클릭 경로는 focus-out 부수효과로 이미
    /// 자동 commit 되지만, 단축키 경로는 focus가 유지되므로 명시적으로
    /// commit 해주어야 합니다.)
    ///
    /// `ON_LOAD` 시 단축키를 자동 등록한 뒤 `on_load`를 호출합니다.
    /// `Self: Sized` 제약은 `set_action_callback`의 단형화에 필요합니다.
    fn dispatch(&self, action_name: &str, hwp: &HwpObject) -> Result<bool>
    where
        Self: Sized,
    {
        match action_name {
            UUIDSTR_ON_INITIAL_LOAD => self.on_initial_load(hwp),
            UUIDSTR_ON_LOAD => {
                crate::shortcut::set_hwp_dispatch(hwp);
                crate::shortcut::set_action_callback(self);
                crate::shortcut::register_action_shortcuts(self);
                self.on_load(hwp)
            }
            _ => {
                crate::ime::commit_composition();
                self.do_action(action_name, hwp)
            }
        }
    }

    /// 예약 GUID(`ON_INITIAL_LOAD`, `ON_LOAD`)는 UI 업데이트 없이 `0`을 반환하고,
    /// 그 외 액션은 [`update_ui`](Self::update_ui)로 라우팅합니다.
    fn dispatch_update_ui(&self, action_name: &str, hwp: &HwpObject) -> u32 {
        match action_name {
            UUIDSTR_ON_INITIAL_LOAD | UUIDSTR_ON_LOAD => 0,
            _ => self.update_ui(action_name, hwp),
        }
    }
}
