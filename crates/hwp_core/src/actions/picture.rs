/// 그림 속성·효과 관련 액션
///
/// SDK 참고: ActionTable_2504.pdf § Picture*
use crate::hwp_obj::HwpObject;

impl HwpObject {
    /// `PictureInsertDialog` — 그림 넣기 대화상자 (API 용)
    pub fn picture_insert_dialog(&self) -> crate::error::Result<()> {
        self.h_action()?.run("PictureInsertDialog")
    }

    /// `PictureChange` — 그림 바꾸기 (ParameterSet: `PictureChange`)
    pub fn picture_change(&self) -> crate::error::Result<()> {
        self.h_action()?.run("PictureChange")
    }

    /// `PictureToOriginal` — 그림 원래 그림으로 되돌리기
    pub fn picture_to_original(&self) -> crate::error::Result<()> {
        self.h_action()?.run("PictureToOriginal")
    }

    /// `PictureLinkedToEmbedded` — 연결된 그림을 모두 삽입그림으로 변환
    pub fn picture_linked_to_embedded(&self) -> crate::error::Result<()> {
        self.h_action()?.run("PictureLinkedToEmbedded")
    }

    /// `PictureSave` — 그림 빼내기
    pub fn picture_save(&self) -> crate::error::Result<()> {
        self.h_action()?.run("PictureSave")
    }

    /// `PictureSaveAsAll` — 삽입된 바이너리 그림 다른 형태로 저장 (ParameterSet: `SaveAsImage`)
    pub fn picture_save_as_all(&self) -> crate::error::Result<()> {
        self.h_action()?.run("PictureSaveAsAll")
    }

    /// `PictureSaveAsOption` — 바이너리 그림을 다른 형태로 저장하는 옵션 설정 (ParameterSet: `SaveAsImage`)
    pub fn picture_save_as_option(&self) -> crate::error::Result<()> {
        self.h_action()?.run("PictureSaveAsOption")
    }

    /// `PictureScissor` — 그림 자르기
    pub fn picture_scissor(&self) -> crate::error::Result<()> {
        self.h_action()?.run("PictureScissor")
    }

    // ── 그림 효과 ──

    /// `PictureEffect1` — 그림 그레이 스케일
    pub fn picture_effect1(&self) -> crate::error::Result<()> {
        self.h_action()?.run("PictureEffect1")
    }

    /// `PictureEffect2` — 그림 흑백으로
    pub fn picture_effect2(&self) -> crate::error::Result<()> {
        self.h_action()?.run("PictureEffect2")
    }

    /// `PictureEffect3` — 그림 워터마크
    pub fn picture_effect3(&self) -> crate::error::Result<()> {
        self.h_action()?.run("PictureEffect3")
    }

    /// `PictureEffect4` — 그림 효과 없음
    pub fn picture_effect4(&self) -> crate::error::Result<()> {
        self.h_action()?.run("PictureEffect4")
    }

    /// `PictureEffect5` — 그림 밝기 증가
    pub fn picture_effect5(&self) -> crate::error::Result<()> {
        self.h_action()?.run("PictureEffect5")
    }

    /// `PictureEffect6` — 그림 밝기 감소
    pub fn picture_effect6(&self) -> crate::error::Result<()> {
        self.h_action()?.run("PictureEffect6")
    }

    /// `PictureEffect7` — 그림 명암 증가
    pub fn picture_effect7(&self) -> crate::error::Result<()> {
        self.h_action()?.run("PictureEffect7")
    }

    /// `PictureEffect8` — 그림 명암 감소
    pub fn picture_effect8(&self) -> crate::error::Result<()> {
        self.h_action()?.run("PictureEffect8")
    }

    // ── 효과 없음 (ShapeObject 기반) ──

    /// `PictureNoBrightness` — 그림 밝기 효과 없음 (ParameterSet: `ShapeObject`)
    pub fn picture_no_brightness(&self) -> crate::error::Result<()> {
        self.h_action()?.run("PictureNoBrightness")
    }

    /// `PictureNoContrast` — 그림 대비 효과 없음 (ParameterSet: `ShapeObject`)
    pub fn picture_no_contrast(&self) -> crate::error::Result<()> {
        self.h_action()?.run("PictureNoContrast")
    }

    /// `PictureNoGlow` — 그림 네온 효과 없음 (ParameterSet: `ShapeObject`)
    pub fn picture_no_glow(&self) -> crate::error::Result<()> {
        self.h_action()?.run("PictureNoGlow")
    }

    /// `PictureNoReflection` — 그림 반사 효과 없음 (ParameterSet: `ShapeObject`)
    pub fn picture_no_reflection(&self) -> crate::error::Result<()> {
        self.h_action()?.run("PictureNoReflection")
    }

    /// `PictureNoShadow` — 그림 그림자 효과 없음 (ParameterSet: `ShapeObject`)
    pub fn picture_no_shadow(&self) -> crate::error::Result<()> {
        self.h_action()?.run("PictureNoShadow")
    }

    /// `PictureNoSofeEdge` — 그림 부드러운 가장자리 효과 없음 (ParameterSet: `ShapeObject`)
    pub fn picture_no_sofe_edge(&self) -> crate::error::Result<()> {
        self.h_action()?.run("PictureNoSofeEdge")
    }

    /// `PictureNoStyle` — 그림 스타일 효과 없음 (ParameterSet: `ShapeObject`)
    pub fn picture_no_style(&self) -> crate::error::Result<()> {
        self.h_action()?.run("PictureNoStyle")
    }

    /// `NoneTextArtShadow` — 글맵시 그림자 없음 (ParameterSet: `ShapeObject`)
    pub fn none_text_art_shadow(&self) -> crate::error::Result<()> {
        self.h_action()?.run("NoneTextArtShadow")
    }
}
