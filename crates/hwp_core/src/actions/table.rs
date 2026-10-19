use crate::method::h_action::HAction;

/// 표(Table) 관련 액션
///
/// SDK 참고: ActionTable_2504.pdf § Table*
impl HAction {
    // ── 표 만들기·합치기·나누기 ──

    /// `TableCreate` — 표 만들기 (ParameterSet: `TableCreation`)
    pub fn table_create(&self) -> crate::error::Result<()> {
        self.run("TableCreate")
    }

    /// `TableMergeTable` — 표 붙이기
    pub fn table_merge_table(&self) -> crate::error::Result<()> {
        self.run("TableMergeTable")
    }

    /// `TableSplitTable` — 표 나누기
    pub fn table_split_table(&self) -> crate::error::Result<()> {
        self.run("TableSplitTable")
    }

    /// `TableSwap` — 표 뒤집기 (ParameterSet: `TableSwap`)
    pub fn table_swap(&self) -> crate::error::Result<()> {
        self.run("TableSwap")
    }

    /// `TableTemplate` — 표 마당 (ParameterSet: `TableTemplate`)
    pub fn table_template(&self) -> crate::error::Result<()> {
        self.run("TableTemplate")
    }

    /// `TablePropertyDialog` — 표 고치기 (ParameterSet: `ShapeObject`)
    pub fn table_property_dialog(&self) -> crate::error::Result<()> {
        self.run("TablePropertyDialog")
    }

    /// `TableTreatAsChar` — 표 글자처럼 취급 (ParameterSet: `ShapeObject`)
    pub fn table_treat_as_char(&self) -> crate::error::Result<()> {
        self.run("TableTreatAsChar")
    }

    // ── 줄·칸 삽입/삭제 ──

    /// `TableAppendRow` — 줄 추가
    pub fn table_append_row(&self) -> crate::error::Result<()> {
        self.run("TableAppendRow")
    }

    /// `TableInsertUpperRow` — 위쪽 줄 삽입 (ParameterSet: `TableInsertLine`)
    pub fn table_insert_upper_row(&self) -> crate::error::Result<()> {
        self.run("TableInsertUpperRow")
    }

    /// `TableInsertLowerRow` — 아래쪽 줄 삽입 (ParameterSet: `TableInsertLine`)
    pub fn table_insert_lower_row(&self) -> crate::error::Result<()> {
        self.run("TableInsertLowerRow")
    }

    /// `TableInsertLeftColumn` — 왼쪽 칸 삽입 (ParameterSet: `TableInsertLine`)
    pub fn table_insert_left_column(&self) -> crate::error::Result<()> {
        self.run("TableInsertLeftColumn")
    }

    /// `TableInsertRightColumn` — 오른쪽 칸 삽입 (ParameterSet: `TableInsertLine`)
    pub fn table_insert_right_column(&self) -> crate::error::Result<()> {
        self.run("TableInsertRightColumn")
    }

    /// `TableInsertRowColumn` — 줄-칸 삽입 (ParameterSet: `TableInsertLine`)
    pub fn table_insert_row_column(&self) -> crate::error::Result<()> {
        self.run("TableInsertRowColumn")
    }

    /// `TableDeleteRow` — 줄 지우기 (ParameterSet: `TableDeleteLine`)
    pub fn table_delete_row(&self) -> crate::error::Result<()> {
        self.run("TableDeleteRow")
    }

    /// `TableDeleteColumn` — 칸 지우기 (ParameterSet: `TableDeleteLine`)
    pub fn table_delete_column(&self) -> crate::error::Result<()> {
        self.run("TableDeleteColumn")
    }

    /// `TableDeleteRowColumn` — 줄-칸 지우기 (ParameterSet: `TableDeleteLine`)
    pub fn table_delete_row_column(&self) -> crate::error::Result<()> {
        self.run("TableDeleteRowColumn")
    }

    /// `TableSubtractRow` — 표 줄 삭제 (ParameterSet: `TableDeleteLine`)
    pub fn table_subtract_row(&self) -> crate::error::Result<()> {
        self.run("TableSubtractRow")
    }

    /// `TableDeleteCell` — 셀 삭제
    pub fn table_delete_cell(&self) -> crate::error::Result<()> {
        self.run("TableDeleteCell")
    }

    // ── 셀 합치기·나누기 ──

    /// `TableMergeCell` — 셀 합치기
    pub fn table_merge_cell(&self) -> crate::error::Result<()> {
        self.run("TableMergeCell")
    }

    /// `TableSplitCell` — 셀 나누기 (ParameterSet: `TableSplitCell`)
    pub fn table_split_cell(&self) -> crate::error::Result<()> {
        self.run("TableSplitCell")
    }

    /// `TableSplitCellCol2` — 셀 칸 나누기 (ParameterSet: `TableSplitCell`)
    pub fn table_split_cell_col2(&self) -> crate::error::Result<()> {
        self.run("TableSplitCellCol2")
    }

    /// `TableSplitCellRow2` — 셀 줄 나누기 (ParameterSet: `TableSplitCell`)
    pub fn table_split_cell_row2(&self) -> crate::error::Result<()> {
        self.run("TableSplitCellRow2")
    }

    // ── 셀 이동 ──

    /// `TableUpperCell` — 셀 이동: 셀 위
    pub fn table_upper_cell(&self) -> crate::error::Result<()> {
        self.run("TableUpperCell")
    }

    /// `TableLowerCell` — 셀 이동: 셀 아래
    pub fn table_lower_cell(&self) -> crate::error::Result<()> {
        self.run("TableLowerCell")
    }

    /// `TableLeftCell` — 셀 이동: 셀 왼쪽
    pub fn table_left_cell(&self) -> crate::error::Result<()> {
        self.run("TableLeftCell")
    }

    /// `TableRightCell` — 셀 이동: 셀 오른쪽
    pub fn table_right_cell(&self) -> crate::error::Result<()> {
        self.run("TableRightCell")
    }

    /// `TableRightCellAppend` — 셀 이동: 셀 오른쪽에 이어서
    pub fn table_right_cell_append(&self) -> crate::error::Result<()> {
        self.run("TableRightCellAppend")
    }

    /// `TableColBegin` — 셀 이동: 열 시작
    pub fn table_col_begin(&self) -> crate::error::Result<()> {
        self.run("TableColBegin")
    }

    /// `TableColEnd` — 셀 이동: 열 끝
    pub fn table_col_end(&self) -> crate::error::Result<()> {
        self.run("TableColEnd")
    }

    /// `TableColPageUp` — 셀 이동: 페이지 업
    pub fn table_col_page_up(&self) -> crate::error::Result<()> {
        self.run("TableColPageUp")
    }

    /// `TableColPageDown` — 셀 이동: 페이지 다운
    pub fn table_col_page_down(&self) -> crate::error::Result<()> {
        self.run("TableColPageDown")
    }

    // ── 셀 블록 선택 ──

    /// `TableCellBlock` — 셀 블록 선택
    pub fn table_cell_block(&self) -> crate::error::Result<()> {
        self.run("TableCellBlock")
    }

    /// `TableCellBlockCol` — 셀 블록 선택 (칸)
    pub fn table_cell_block_col(&self) -> crate::error::Result<()> {
        self.run("TableCellBlockCol")
    }

    /// `TableCellBlockRow` — 셀 블록 선택 (줄)
    pub fn table_cell_block_row(&self) -> crate::error::Result<()> {
        self.run("TableCellBlockRow")
    }

    /// `TableCellBlockExtend` — 셀 블록 연장 (F5 + F5)
    pub fn table_cell_block_extend(&self) -> crate::error::Result<()> {
        self.run("TableCellBlockExtend")
    }

    /// `TableCellBlockExtendAbs` — 셀 블록 연장 (SHIFT + F5)
    pub fn table_cell_block_extend_abs(&self) -> crate::error::Result<()> {
        self.run("TableCellBlockExtendAbs")
    }

    // ── 셀 크기 조정 ──

    /// `TableResizeCellDown` — 셀 크기 변경: 셀 아래
    pub fn table_resize_cell_down(&self) -> crate::error::Result<()> {
        self.run("TableResizeCellDown")
    }

    /// `TableResizeCellUp` — 셀 크기 변경: 셀 위
    pub fn table_resize_cell_up(&self) -> crate::error::Result<()> {
        self.run("TableResizeCellUp")
    }

    /// `TableResizeCellLeft` — 셀 크기 변경: 셀 왼쪽
    pub fn table_resize_cell_left(&self) -> crate::error::Result<()> {
        self.run("TableResizeCellLeft")
    }

    /// `TableResizeCellRight` — 셀 크기 변경: 셀 오른쪽
    pub fn table_resize_cell_right(&self) -> crate::error::Result<()> {
        self.run("TableResizeCellRight")
    }

    /// `TableResizeDown` — 셀 크기 변경 (아래)
    pub fn table_resize_down(&self) -> crate::error::Result<()> {
        self.run("TableResizeDown")
    }

    /// `TableResizeUp` — 셀 크기 변경 (위)
    pub fn table_resize_up(&self) -> crate::error::Result<()> {
        self.run("TableResizeUp")
    }

    /// `TableResizeLeft` — 셀 크기 변경 (왼쪽)
    pub fn table_resize_left(&self) -> crate::error::Result<()> {
        self.run("TableResizeLeft")
    }

    /// `TableResizeRight` — 셀 크기 변경 (오른쪽)
    pub fn table_resize_right(&self) -> crate::error::Result<()> {
        self.run("TableResizeRight")
    }

    /// `TableResizeLineDown` — 셀 크기 변경: 선아래
    pub fn table_resize_line_down(&self) -> crate::error::Result<()> {
        self.run("TableResizeLineDown")
    }

    /// `TableResizeLineUp` — 셀 크기 변경: 선 위
    pub fn table_resize_line_up(&self) -> crate::error::Result<()> {
        self.run("TableResizeLineUp")
    }

    /// `TableResizeLineLeft` — 셀 크기 변경: 선 왼쪽
    pub fn table_resize_line_left(&self) -> crate::error::Result<()> {
        self.run("TableResizeLineLeft")
    }

    /// `TableResizeLineRight` — 셀 크기 변경: 선 오른쪽
    pub fn table_resize_line_right(&self) -> crate::error::Result<()> {
        self.run("TableResizeLineRight")
    }

    /// `TableResizeExDown` — 셀 크기 변경: 셀 아래 (셀 블록 상태 아니어도 동작)
    pub fn table_resize_ex_down(&self) -> crate::error::Result<()> {
        self.run("TableResizeExDown")
    }

    /// `TableResizeExUp` — 셀 크기 변경: 셀 위 (셀 블록 상태 아니어도 동작)
    pub fn table_resize_ex_up(&self) -> crate::error::Result<()> {
        self.run("TableResizeExUp")
    }

    /// `TableResizeExLeft` — 셀 크기 변경: 셀 왼쪽 (셀 블록 상태 아니어도 동작)
    pub fn table_resize_ex_left(&self) -> crate::error::Result<()> {
        self.run("TableResizeExLeft")
    }

    /// `TableResizeExRight` — 셀 크기 변경: 셀 오른쪽 (셀 블록 상태 아니어도 동작)
    pub fn table_resize_ex_right(&self) -> crate::error::Result<()> {
        self.run("TableResizeExRight")
    }

    /// `TableDistributeCellHeight` — 셀 높이를 같게
    pub fn table_distribute_cell_height(&self) -> crate::error::Result<()> {
        self.run("TableDistributeCellHeight")
    }

    /// `TableDistributeCellWidth` — 셀 너비를 같게
    pub fn table_distribute_cell_width(&self) -> crate::error::Result<()> {
        self.run("TableDistributeCellWidth")
    }

    // ── 셀 테두리 ──

    /// `TableCellBorderAll` — 모든 셀 테두리 toggle (셀 블록 상태에서만)
    pub fn table_cell_border_all(&self) -> crate::error::Result<()> {
        self.run("TableCellBorderAll")
    }

    /// `TableCellBorderNo` — 모든 셀 테두리 지움 (셀 블록 상태에서만)
    pub fn table_cell_border_no(&self) -> crate::error::Result<()> {
        self.run("TableCellBorderNo")
    }

    /// `TableCellBorderOutside` — 바깥 셀 테두리 toggle (셀 블록 상태에서만)
    pub fn table_cell_border_outside(&self) -> crate::error::Result<()> {
        self.run("TableCellBorderOutside")
    }

    /// `TableCellBorderInside` — 모든 안쪽 셀 테두리 toggle (셀 블록 상태에서만)
    pub fn table_cell_border_inside(&self) -> crate::error::Result<()> {
        self.run("TableCellBorderInside")
    }

    /// `TableCellBorderInsideHorz` — 모든 안쪽 가로 셀 테두리 toggle (셀 블록 상태에서만)
    pub fn table_cell_border_inside_horz(&self) -> crate::error::Result<()> {
        self.run("TableCellBorderInsideHorz")
    }

    /// `TableCellBorderInsideVert` — 모든 안쪽 세로 셀 테두리 toggle (셀 블록 상태에서만)
    pub fn table_cell_border_inside_vert(&self) -> crate::error::Result<()> {
        self.run("TableCellBorderInsideVert")
    }

    /// `TableCellBorderTop` — 가장 위의 셀 테두리 toggle (셀 블록 상태에서만)
    pub fn table_cell_border_top(&self) -> crate::error::Result<()> {
        self.run("TableCellBorderTop")
    }

    /// `TableCellBorderBottom` — 가장 아래 셀 테두리 toggle (셀 블록 상태에서만)
    pub fn table_cell_border_bottom(&self) -> crate::error::Result<()> {
        self.run("TableCellBorderBottom")
    }

    /// `TableCellBorderLeft` — 가장 왼쪽의 셀 테두리 toggle (셀 블록 상태에서만)
    pub fn table_cell_border_left(&self) -> crate::error::Result<()> {
        self.run("TableCellBorderLeft")
    }

    /// `TableCellBorderRight` — 가장 오른쪽의 셀 테두리 toggle (셀 블록 상태에서만)
    pub fn table_cell_border_right(&self) -> crate::error::Result<()> {
        self.run("TableCellBorderRight")
    }

    /// `TableCellBorderDiagonalDown` — 대각선(↘) 셀 테두리 toggle (셀 블록 상태에서만)
    pub fn table_cell_border_diagonal_down(&self) -> crate::error::Result<()> {
        self.run("TableCellBorderDiagonalDown")
    }

    /// `TableCellBorderDiagonalUp` — 대각선(↗) 셀 테두리 toggle (셀 블록 상태에서만)
    pub fn table_cell_border_diagonal_up(&self) -> crate::error::Result<()> {
        self.run("TableCellBorderDiagonalUp")
    }

    // ── 셀 음영·방향 ──

    /// `TableCellShadeDec` — 셀 배경의 음영을 낮춥니다 (ParameterSet: `CellBorderFill`)
    pub fn table_cell_shade_dec(&self) -> crate::error::Result<()> {
        self.run("TableCellShadeDec")
    }

    /// `TableCellShadeInc` — 셀 배경의 음영을 높입니다 (ParameterSet: `CellBorderFill`)
    pub fn table_cell_shade_inc(&self) -> crate::error::Result<()> {
        self.run("TableCellShadeInc")
    }

    /// `TableCellTextHorz` — 셀 문자 방향-가로 쓰기 (ParameterSet: `CellBorderFill`)
    pub fn table_cell_text_horz(&self) -> crate::error::Result<()> {
        self.run("TableCellTextHorz")
    }

    /// `TableCellTextVert` — 셀 음열 없음 (ParameterSet: `CellBorderFill`)
    pub fn table_cell_text_vert(&self) -> crate::error::Result<()> {
        self.run("TableCellTextVert")
    }

    /// `TableCellTextVertAll` — 셀 문자 방향-세로 쓰기-영문 세움 (ParameterSet: `CellBorderFill`)
    pub fn table_cell_text_vert_all(&self) -> crate::error::Result<()> {
        self.run("TableCellTextVertAll")
    }

    /// `TableCellToggleDirection` — 표 문자 방향 toggle (ParameterSet: `CellBorderFill`)
    pub fn table_cell_toggle_direction(&self) -> crate::error::Result<()> {
        self.run("TableCellToggleDirection")
    }

    // ── 셀 정렬 ──

    /// `TableVAlignTop` — 셀 세로정렬: 위
    pub fn table_v_align_top(&self) -> crate::error::Result<()> {
        self.run("TableVAlignTop")
    }

    /// `TableVAlignCenter` — 셀 세로정렬: 가운데
    pub fn table_v_align_center(&self) -> crate::error::Result<()> {
        self.run("TableVAlignCenter")
    }

    /// `TableVAlignBottom` — 셀 세로정렬: 아래
    pub fn table_v_align_bottom(&self) -> crate::error::Result<()> {
        self.run("TableVAlignBottom")
    }

    // ── 쪽 경계 ──

    /// `TableBreak` — 표 쪽 경계에서 (나누지 않음, ParameterSet: `Table`)
    pub fn table_break(&self) -> crate::error::Result<()> {
        self.run("TableBreak")
    }

    /// `TableBreakCell` — 표 쪽 경계에서 (나눔, ParameterSet: `Table`)
    pub fn table_break_cell(&self) -> crate::error::Result<()> {
        self.run("TableBreakCell")
    }

    /// `TableBreakNone` — 표 쪽 경계에서 (셀 단위로 나눔, ParameterSet: `Table`)
    pub fn table_break_none(&self) -> crate::error::Result<()> {
        self.run("TableBreakNone")
    }

    // ── 자동 채우기 ──

    /// `TableAutoFill` — 자동 채우기 (ParameterSet: `AutoFill*`)
    pub fn table_auto_fill(&self) -> crate::error::Result<()> {
        self.run("TableAutoFill")
    }

    /// `TableAutoFillDlg` — 자동 채우기 대화상자 (ParameterSet: `AutoFill*`)
    pub fn table_auto_fill_dlg(&self) -> crate::error::Result<()> {
        self.run("TableAutoFillDlg")
    }

    // ── 계산식 ──

    /// `TableFormula` — 계산식 (ParameterSet: `FieldCtrl`)
    pub fn table_formula(&self) -> crate::error::Result<()> {
        self.run("TableFormula")
    }

    /// `TableFormulaSumAuto` — 블록 합계
    pub fn table_formula_sum_auto(&self) -> crate::error::Result<()> {
        self.run("TableFormulaSumAuto")
    }

    /// `TableFormulaSumHor` — 가로 합계
    pub fn table_formula_sum_hor(&self) -> crate::error::Result<()> {
        self.run("TableFormulaSumHor")
    }

    /// `TableFormulaSumVer` — 세로 합계
    pub fn table_formula_sum_ver(&self) -> crate::error::Result<()> {
        self.run("TableFormulaSumVer")
    }

    /// `TableFormulaAvgAuto` — 블록 평균
    pub fn table_formula_avg_auto(&self) -> crate::error::Result<()> {
        self.run("TableFormulaAvgAuto")
    }

    /// `TableFormulaAvgHor` — 가로 평균
    pub fn table_formula_avg_hor(&self) -> crate::error::Result<()> {
        self.run("TableFormulaAvgHor")
    }

    /// `TableFormulaAvgVer` — 세로 평균
    pub fn table_formula_avg_ver(&self) -> crate::error::Result<()> {
        self.run("TableFormulaAvgVer")
    }

    /// `TableFormulaProAuto` — 블록 곱
    pub fn table_formula_pro_auto(&self) -> crate::error::Result<()> {
        self.run("TableFormulaProAuto")
    }

    /// `TableFormulaProHor` — 가로 곱
    pub fn table_formula_pro_hor(&self) -> crate::error::Result<()> {
        self.run("TableFormulaProHor")
    }

    /// `TableFormulaProVer` — 세로 곱
    pub fn table_formula_pro_ver(&self) -> crate::error::Result<()> {
        self.run("TableFormulaProVer")
    }

    // ── 문자열 변환 ──

    /// `TableStringToTable` — 문자열을 표로 변환 (ParameterSet: `TableStrToTbl`)
    pub fn table_string_to_table(&self) -> crate::error::Result<()> {
        self.run("TableStringToTable")
    }

    /// `TableTableToString` — 표를 문자열로 변환 (ParameterSet: `TableTblToStr`)
    pub fn table_table_to_string(&self) -> crate::error::Result<()> {
        self.run("TableTableToString")
    }

    // ── 세자리 구분 ──

    /// `TableInsertComma` — 세자리마다 자리점 넣기
    pub fn table_insert_comma(&self) -> crate::error::Result<()> {
        self.run("TableInsertComma")
    }

    /// `TableDeleteComma` — 세자리마다 자리점 빼기
    pub fn table_delete_comma(&self) -> crate::error::Result<()> {
        self.run("TableDeleteComma")
    }

    // ── 그리기 펜·지우개 ──

    /// `TableDrawPen` — 표 그리기
    pub fn table_draw_pen(&self) -> crate::error::Result<()> {
        self.run("TableDrawPen")
    }

    /// `TableAutoDrawPenStyleWidthDlg` — 표 그리기 선 모양 (ParameterSet: `TableDrawPen`)
    pub fn table_auto_draw_pen_style_width_dlg(&self) -> crate::error::Result<()> {
        self.run("TableAutoDrawPenStyleWidthDlg")
    }

    /// `TableEraser` — 표 지우개
    pub fn table_eraser(&self) -> crate::error::Result<()> {
        self.run("TableEraser")
    }

    // ── 캡션 위치 ──

    /// `TableCaptionPosTop` — 테이블 캡션 위치-위
    pub fn table_caption_pos_top(&self) -> crate::error::Result<()> {
        self.run("TableCaptionPosTop")
    }

    /// `TableCaptionPosBottom` — 테이블 캡션 위치-아래
    pub fn table_caption_pos_bottom(&self) -> crate::error::Result<()> {
        self.run("TableCaptionPosBottom")
    }

    /// `TableCaptionPosLeftTop` — 테이블 캡션 위치-왼쪽 위
    pub fn table_caption_pos_left_top(&self) -> crate::error::Result<()> {
        self.run("TableCaptionPosLeftTop")
    }

    /// `TableCaptionPosLeftCenter` — 테이블 캡션 위치-왼쪽 가운데
    pub fn table_caption_pos_left_center(&self) -> crate::error::Result<()> {
        self.run("TableCaptionPosLeftCenter")
    }

    /// `TableCaptionPosLeftBottom` — 테이블 캡션 위치-왼쪽 아래
    pub fn table_caption_pos_left_bottom(&self) -> crate::error::Result<()> {
        self.run("TableCaptionPosLeftBottom")
    }

    /// `TableCaptionPosRightTop` — 테이블 캡션 위치-오른쪽 위
    pub fn table_caption_pos_right_top(&self) -> crate::error::Result<()> {
        self.run("TableCaptionPosRightTop")
    }

    /// `TableCaptionPosRightCenter` — 테이블 캡션 위치-오른쪽 가운데
    pub fn table_caption_pos_right_center(&self) -> crate::error::Result<()> {
        self.run("TableCaptionPosRightCenter")
    }

    /// `TableCaptionPosRightBottom` — 테이블 캡션 위치-오른쪽 아래
    pub fn table_caption_pos_right_bottom(&self) -> crate::error::Result<()> {
        self.run("TableCaptionPosRightBottom")
    }

    // ── 셀 정렬 (가로+세로 조합) ──

    /// `TableCellAlignLeftTop` — para align + cell valign (테이블 셀)
    pub fn table_cell_align_left_top(&self) -> crate::error::Result<()> {
        self.run("TableCellAlignLeftTop")
    }

    /// `TableCellAlignLeftCenter` — para align + cell valign (테이블 셀)
    pub fn table_cell_align_left_center(&self) -> crate::error::Result<()> {
        self.run("TableCellAlignLeftCenter")
    }

    /// `TableCellAlignLeftBottom` — para align + cell valign (테이블 셀)
    pub fn table_cell_align_left_bottom(&self) -> crate::error::Result<()> {
        self.run("TableCellAlignLeftBottom")
    }

    /// `TableCellAlignCenterTop` — para align + cell valign (테이블 셀)
    pub fn table_cell_align_center_top(&self) -> crate::error::Result<()> {
        self.run("TableCellAlignCenterTop")
    }

    /// `TableCellAlignCenterCenter` — para align + cell valign (테이블 셀)
    pub fn table_cell_align_center_center(&self) -> crate::error::Result<()> {
        self.run("TableCellAlignCenterCenter")
    }

    /// `TableCellAlignCenterBottom` — para align + cell valign (테이블 셀)
    pub fn table_cell_align_center_bottom(&self) -> crate::error::Result<()> {
        self.run("TableCellAlignCenterBottom")
    }

    /// `TableCellAlignRightTop` — para align + cell valign (테이블 셀)
    pub fn table_cell_align_right_top(&self) -> crate::error::Result<()> {
        self.run("TableCellAlignRightTop")
    }

    /// `TableCellAlignRightCenter` — para align + cell valign (테이블 셀)
    pub fn table_cell_align_right_center(&self) -> crate::error::Result<()> {
        self.run("TableCellAlignRightCenter")
    }

    /// `TableCellAlignRightBottom` — para align + cell valign (테이블 셀)
    pub fn table_cell_align_right_bottom(&self) -> crate::error::Result<()> {
        self.run("TableCellAlignRightBottom")
    }
}
