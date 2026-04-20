/// 글자 모양 관련 액션
///
/// SDK 참고: ActionTable_2504.pdf § CharShape*
///
/// `CharShapeHeight`, `CharShapeTextColor` 등 파라미터(`CharShape*`)가 필요한 액션은
/// `HAction.GetDefault("CharShape", hset) + Execute` 패턴으로 값을 설정하십시오.
use crate::hwp_obj::HwpObject;

impl HwpObject {
    /// `CharShapeBold` — 굵게 토글
    pub fn char_shape_bold(&self) -> crate::error::Result<()> {
        self.h_action()?.run("CharShapeBold")
    }

    /// `CharShapeItalic` — 기울임꼴 토글
    pub fn char_shape_italic(&self) -> crate::error::Result<()> {
        self.h_action()?.run("CharShapeItalic")
    }

    /// `CharShapeUnderline` — 밑줄 토글
    pub fn char_shape_underline(&self) -> crate::error::Result<()> {
        self.h_action()?.run("CharShapeUnderline")
    }

    /// `CharShapeEmboss` — 양각 토글
    pub fn char_shape_emboss(&self) -> crate::error::Result<()> {
        self.h_action()?.run("CharShapeEmboss")
    }

    /// `CharShapeEngrave` — 음각 토글
    pub fn char_shape_engrave(&self) -> crate::error::Result<()> {
        self.h_action()?.run("CharShapeEngrave")
    }

    /// `CharShapeOutline` — 외곽선 토글
    pub fn char_shape_outline(&self) -> crate::error::Result<()> {
        self.h_action()?.run("CharShapeOutline")
    }

    /// `CharShapeShadow` — 그림자 토글
    pub fn char_shape_shadow(&self) -> crate::error::Result<()> {
        self.h_action()?.run("CharShapeShadow")
    }

    /// `CharShapeSubscript` — 아래 첨자 토글
    pub fn char_shape_subscript(&self) -> crate::error::Result<()> {
        self.h_action()?.run("CharShapeSubscript")
    }

    /// `CharShapeSuperscript` — 위 첨자 토글
    pub fn char_shape_superscript(&self) -> crate::error::Result<()> {
        self.h_action()?.run("CharShapeSuperscript")
    }

    /// `CharShapeNormal` — 글자 모양 초기화 (일반체)
    pub fn char_shape_normal(&self) -> crate::error::Result<()> {
        self.h_action()?.run("CharShapeNormal")
    }

    /// `CharShapeHeight` — 글자 크기 변경 (ParameterSet: `CharShape*`)
    ///
    /// 파라미터 없이 호출하면 기본 동작을 수행합니다.
    /// 크기를 지정하려면 `HAction.GetDefault("CharShape")` 후 `Height` 항목을 설정하십시오.
    pub fn char_shape_height(&self) -> crate::error::Result<()> {
        self.h_action()?.run("CharShapeHeight")
    }

    /// `CharShapeTextColor` — 글자 색 변경 (ParameterSet: `CharShape*`)
    ///
    /// 색을 지정하려면 `HAction.GetDefault("CharShape")` 후 `TextColor` 항목을 설정하십시오.
    pub fn char_shape_text_color(&self) -> crate::error::Result<()> {
        self.h_action()?.run("CharShapeTextColor")
    }

    /// `StyleClearCharStyle` — 글자 스타일을 해제합니다.
    pub fn style_clear_char_style(&self) -> crate::error::Result<()> {
        self.h_action()?.run("StyleClearCharStyle")
    }
}
