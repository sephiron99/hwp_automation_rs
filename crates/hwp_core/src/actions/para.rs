/// 문단 모양·스타일·글머리표·들여쓰기 관련 액션
///
/// SDK 참고: ActionTable_2504.pdf § ParagraphShape*, ParaNumber*, Style* 등
use crate::hwp_obj::HwpObject;

impl HwpObject {
    // ── 문단 모양 대화상자 ──

    /// `ParagraphShape` — 문단 모양 대화상자 (ParameterSet: `ParaShape`)
    pub fn paragraph_shape(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ParagraphShape")
    }

    /// `ParaShapeDialog` — 문단 모양 대화상자 (내부 구현용, ParameterSet: `ParaShape`)
    pub fn para_shape_dialog(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ParaShapeDialog")
    }

    // ── 정렬 ──

    /// `ParagraphShapeAlignLeft` — 왼쪽 정렬
    pub fn paragraph_shape_align_left(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ParagraphShapeAlignLeft")
    }

    /// `ParagraphShapeAlignCenter` — 가운데 정렬
    pub fn paragraph_shape_align_center(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ParagraphShapeAlignCenter")
    }

    /// `ParagraphShapeAlignRight` — 오른쪽 정렬
    pub fn paragraph_shape_align_right(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ParagraphShapeAlignRight")
    }

    /// `ParagraphShapeAlignJustify` — 양쪽 정렬
    pub fn paragraph_shape_align_justify(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ParagraphShapeAlignJustify")
    }

    /// `ParagraphShapeAlignDistribute` — 배분 정렬
    pub fn paragraph_shape_align_distribute(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ParagraphShapeAlignDistribute")
    }

    /// `ParagraphShapeAlignDivision` — 나눔 정렬
    pub fn paragraph_shape_align_division(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ParagraphShapeAlignDivision")
    }

    // ── 여백 ──

    /// `ParagraphShapeDecreaseLeftMargin` — 왼쪽 여백 줄이기
    pub fn paragraph_shape_decrease_left_margin(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ParagraphShapeDecreaseLeftMargin")
    }

    /// `ParagraphShapeIncreaseLeftMargin` — 왼쪽 여백 키우기
    pub fn paragraph_shape_increase_left_margin(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ParagraphShapeIncreaseLeftMargin")
    }

    /// `ParagraphShapeDecreaseRightMargin` — 오른쪽 여백 키우기
    pub fn paragraph_shape_decrease_right_margin(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ParagraphShapeDecreaseRightMargin")
    }

    /// `ParagraphShapeIncreaseRightMargin` — 오른쪽 여백 줄이기
    pub fn paragraph_shape_increase_right_margin(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ParagraphShapeIncreaseRightMargin")
    }

    /// `ParagraphShapeDecreaseMargin` — 왼쪽-오른쪽 여백 줄이기
    pub fn paragraph_shape_decrease_margin(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ParagraphShapeDecreaseMargin")
    }

    /// `ParagraphShapeIncreaseMargin` — 왼쪽-오른쪽 여백 키우기
    pub fn paragraph_shape_increase_margin(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ParagraphShapeIncreaseMargin")
    }

    // ── 줄 간격 ──

    /// `ParagraphShapeDecreaseLineSpacing` — 줄 간격을 점점 줄임
    pub fn paragraph_shape_decrease_line_spacing(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ParagraphShapeDecreaseLineSpacing")
    }

    /// `ParagraphShapeIncreaseLineSpacing` — 줄 간격을 점점 넓힘
    pub fn paragraph_shape_increase_line_spacing(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ParagraphShapeIncreaseLineSpacing")
    }

    // ── 들여쓰기 ──

    /// `ParagraphShapeIndentAtCaret` — 첫 줄 내어 쓰기 (캐럿 위치에서)
    pub fn paragraph_shape_indent_at_caret(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ParagraphShapeIndentAtCaret")
    }

    /// `ParagraphShapeIndentNegative` — 첫 줄을 한 글자 내어 씀
    pub fn paragraph_shape_indent_negative(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ParagraphShapeIndentNegative")
    }

    /// `ParagraphShapeIndentPositive` — 첫 줄을 한 글자 들여 씀
    pub fn paragraph_shape_indent_positive(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ParagraphShapeIndentPositive")
    }

    /// `IndentBlock` — 블록 들여쓰기
    pub fn indent_block(&self) -> crate::error::Result<()> {
        self.h_action()?.run("IndentBlock")
    }

    /// `IndentBlockFixed` — 블록 들여쓰기 (고정)
    pub fn indent_block_fixed(&self) -> crate::error::Result<()> {
        self.h_action()?.run("IndentBlockFixed")
    }

    // ── 기타 문단 속성 ──

    /// `ParagraphShapeProtect` — 문단 보호 토글
    pub fn paragraph_shape_protect(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ParagraphShapeProtect")
    }

    /// `ParagraphShapeSingleRow` — 한 줄로 입력 토글
    pub fn paragraph_shape_single_row(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ParagraphShapeSingleRow")
    }

    /// `ParagraphShapeWithNext` — 다음 문단과 함께 토글
    pub fn paragraph_shape_with_next(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ParagraphShapeWithNext")
    }

    // ── 글머리표·문단번호 ──

    /// `BulletDlg` — 글머리표/문단번호 대화상자 (ParameterSet: `ParaShape`)
    pub fn bullet_dlg(&self) -> crate::error::Result<()> {
        self.h_action()?.run("BulletDlg")
    }

    /// `ParaNumberDlg` — 문단번호 대화상자 (ParameterSet: `ParaShape`)
    pub fn para_number_dlg(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ParaNumberDlg")
    }

    /// `ParaNumberBullet` — 문단번호/글머리표 토글 (ParameterSet: `ParaShape`)
    pub fn para_number_bullet(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ParaNumberBullet")
    }

    /// `ParaNumberBulletLevelDown` — 문단번호/글머리표 한 수준 아래로
    pub fn para_number_bullet_level_down(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ParaNumberBulletLevelDown")
    }

    /// `ParaNumberBulletLevelUp` — 문단번호/글머리표 한 수준 위로
    pub fn para_number_bullet_level_up(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ParaNumberBulletLevelUp")
    }

    /// `PutBullet` — 글머리표 달기 (ParameterSet: `ParaShape*`)
    pub fn put_bullet(&self) -> crate::error::Result<()> {
        self.h_action()?.run("PutBullet")
    }

    /// `PutParaNumber` — 문단번호 달기 (ParameterSet: `ParaShape*`)
    pub fn put_para_number(&self) -> crate::error::Result<()> {
        self.h_action()?.run("PutParaNumber")
    }

    /// `PutNewParaNumber` — 문단번호 새 번호 시작하기 (ParameterSet: `ParaShape*`)
    pub fn put_new_para_number(&self) -> crate::error::Result<()> {
        self.h_action()?.run("PutNewParaNumber")
    }

    /// `PutOutlineNumber` — 개요번호 달기 (ParameterSet: `ParaShape*`)
    pub fn put_outline_number(&self) -> crate::error::Result<()> {
        self.h_action()?.run("PutOutlineNumber")
    }

    /// `PictureBulletDlg` — 그림 글머리표 대화상자 (ParameterSet: `ParaShape`)
    pub fn picture_bullet_dlg(&self) -> crate::error::Result<()> {
        self.h_action()?.run("PictureBulletDlg")
    }

    // ── 내어쓰기 ──

    /// `DropCap` — 문단 첫 글자 장식 (드롭캡) 대화상자 (ParameterSet: `DropCap`)
    pub fn drop_cap(&self) -> crate::error::Result<()> {
        self.h_action()?.run("DropCap")
    }

    // ── 스타일 ──

    /// `Style` — 스타일 대화상자 (ParameterSet: `Style`)
    pub fn style(&self) -> crate::error::Result<()> {
        self.h_action()?.run("Style")
    }

    /// `StyleEx` — 스타일 대화상자 (한글 2007, ParameterSet: `Style`)
    pub fn style_ex(&self) -> crate::error::Result<()> {
        self.h_action()?.run("StyleEx")
    }

    /// `StyleAdd` — 스타일 추가 대화상자 (ParameterSet: `Style`)
    pub fn style_add(&self) -> crate::error::Result<()> {
        self.h_action()?.run("StyleAdd")
    }

    /// `StyleEdit` — 스타일 편집 대화상자 (ParameterSet: `Style`)
    pub fn style_edit(&self) -> crate::error::Result<()> {
        self.h_action()?.run("StyleEdit")
    }

    /// `StyleDelete` — 스타일 제거 (ParameterSet: `StyleDelete`)
    pub fn style_delete(&self) -> crate::error::Result<()> {
        self.h_action()?.run("StyleDelete")
    }

    /// `StyleChangeToCurrentShape` — 스타일을 현재 모양으로 바꾸기 (ParameterSet: `StyleItem`)
    pub fn style_change_to_current_shape(&self) -> crate::error::Result<()> {
        self.h_action()?.run("StyleChangeToCurrentShape")
    }

    /// `StyleParaNumberBullet` — 문단번호/글머리표 스타일 (ParameterSet: `ParaShape`)
    pub fn style_para_number_bullet(&self) -> crate::error::Result<()> {
        self.h_action()?.run("StyleParaNumberBullet")
    }

    /// `StyleTemplate` — 스타일 마당 (ParameterSet: `StyleTemplate`)
    pub fn style_template(&self) -> crate::error::Result<()> {
        self.h_action()?.run("StyleTemplate")
    }

    /// `StyleShortcut1` — 스타일 단축키 `<Ctrl + 1>`
    pub fn style_shortcut1(&self) -> crate::error::Result<()> {
        self.h_action()?.run("StyleShortcut1")
    }

    /// `StyleShortcut2` — 스타일 단축키 `<Ctrl + 2>`
    pub fn style_shortcut2(&self) -> crate::error::Result<()> {
        self.h_action()?.run("StyleShortcut2")
    }

    /// `StyleShortcut3` — 스타일 단축키 `<Ctrl + 3>`
    pub fn style_shortcut3(&self) -> crate::error::Result<()> {
        self.h_action()?.run("StyleShortcut3")
    }

    /// `StyleShortcut4` — 스타일 단축키 `<Ctrl + 4>`
    pub fn style_shortcut4(&self) -> crate::error::Result<()> {
        self.h_action()?.run("StyleShortcut4")
    }

    /// `StyleShortcut5` — 스타일 단축키 `<Ctrl + 5>`
    pub fn style_shortcut5(&self) -> crate::error::Result<()> {
        self.h_action()?.run("StyleShortcut5")
    }

    /// `StyleShortcut6` — 스타일 단축키 `<Ctrl + 6>`
    pub fn style_shortcut6(&self) -> crate::error::Result<()> {
        self.h_action()?.run("StyleShortcut6")
    }

    /// `StyleShortcut7` — 스타일 단축키 `<Ctrl + 7>`
    pub fn style_shortcut7(&self) -> crate::error::Result<()> {
        self.h_action()?.run("StyleShortcut7")
    }

    /// `StyleShortcut8` — 스타일 단축키 `<Ctrl + 8>`
    pub fn style_shortcut8(&self) -> crate::error::Result<()> {
        self.h_action()?.run("StyleShortcut8")
    }

    /// `StyleShortcut9` — 스타일 단축키 `<Ctrl + 9>`
    pub fn style_shortcut9(&self) -> crate::error::Result<()> {
        self.h_action()?.run("StyleShortcut9")
    }

    /// `StyleShortcut10` — 스타일 단축키 `<Ctrl + 0>`
    pub fn style_shortcut10(&self) -> crate::error::Result<()> {
        self.h_action()?.run("StyleShortcut10")
    }
}
