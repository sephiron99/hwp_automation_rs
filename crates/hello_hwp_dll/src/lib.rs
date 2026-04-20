use hwp_addon::debug::msgbox_err;
use hwp_addon::export_hwp_addon;
use hwp_addon::hwp_user_action::{ActionMeta, HwpUserAction, ToolbarConfig, ToolbarTarget};
use hwp_core::hwp_obj::HwpObject;

const TOOLBAR_DATA: &[u8] = include_bytes!("../toolbar.bmp");

const MY_ACTION: &str = "InsertHelloWorld";

static CONFIG: ToolbarConfig = ToolbarConfig {
    name: "addon",
    serialize_path: "HelloHwpDll",
    bitmap_data: TOOLBAR_DATA,
    target: ToolbarTarget::Both("추가 기능"),
    ribbon_toolbox_index: -1,
};

pub struct HelloWorldPlugin;

impl HwpUserAction for HelloWorldPlugin {
    fn toolbar_config(&self) -> Option<&ToolbarConfig> {
        Some(&CONFIG)
    }

    fn actions(&self) -> &'static [ActionMeta] {
        static ACTIONS: [ActionMeta; 1] = [ActionMeta {
            name: MY_ACTION,
            label: "Hello",
            image_index: 0,
            shortcut: None,
        }];
        &ACTIONS
    }

    fn on_load(&self, hwp: &HwpObject) -> hwp_core::error::Result<bool> {
        match self.setup_toolbar(hwp) {
            Ok(v) => Ok(v),
            Err(e) => {
                msgbox_err("on_load", &format!("{e}"));
                Err(e)
            }
        }
    }

    fn do_action(&self, action_name: &str, hwp: &HwpObject) -> hwp_core::error::Result<bool> {
        match action_name {
            MY_ACTION => {
                let action = hwp.h_action()?;
                let pset = hwp.h_parameter_set()?;
                let insert_text = pset.h_insert_text()?;
                let hset = insert_text.h_set()?;

                action.get_default("InsertText", &hset)?;
                insert_text.set_text("Hello World! Rust 애드인에서 삽입된 텍스트입니다.")?;
                action.execute("InsertText", &hset)?;

                // hwp.move_pos(hwp_core::ihwpobject::movepos::MovePos::StartOfPara)?;
                // must return Err
                // action.run("WrongAction")?;

                Ok(true)
            }
            _ => Ok(false),
        }
    }
}

export_hwp_addon!(HelloWorldPlugin, HelloWorldPlugin);
