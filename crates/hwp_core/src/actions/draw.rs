use crate::method::h_action::HAction;

/// 그리기 개체·도형·글상자 관련 액션
///
/// SDK 참고: ActionTable_2504.pdf § DrawObjCreator*, ShapeObj*, TextArt*, TextBox*
impl HAction {
    // ── 도형 생성 ──

    /// `DrawObjCreatorLine` — 선 그리기
    pub fn draw_obj_creator_line(&self) -> crate::error::Result<()> {
        self.run("DrawObjCreatorLine")
    }

    /// `DrawObjCreatorRect` — 직사각형 그리기
    pub fn draw_obj_creator_rect(&self) -> crate::error::Result<()> {
        self.run("DrawObjCreatorRect")
    }

    /// `DrawObjCreatorRoundRect` — 둥근 직사각형 그리기
    pub fn draw_obj_creator_round_rect(&self) -> crate::error::Result<()> {
        self.run("DrawObjCreatorRoundRect")
    }

    /// `DrawObjCreatorEllipse` — 타원 그리기
    pub fn draw_obj_creator_ellipse(&self) -> crate::error::Result<()> {
        self.run("DrawObjCreatorEllipse")
    }

    /// `DrawObjCreatorArc` — 호 그리기
    pub fn draw_obj_creator_arc(&self) -> crate::error::Result<()> {
        self.run("DrawObjCreatorArc")
    }

    /// `DrawObjCreatorMultiArc` — 다각호 그리기
    pub fn draw_obj_creator_multi_arc(&self) -> crate::error::Result<()> {
        self.run("DrawObjCreatorMultiArc")
    }

    /// `DrawObjCreatorCurve` — 곡선 그리기
    pub fn draw_obj_creator_curve(&self) -> crate::error::Result<()> {
        self.run("DrawObjCreatorCurve")
    }

    /// `DrawObjCreatorMultiCurve` — 다각 곡선 그리기
    pub fn draw_obj_creator_multi_curve(&self) -> crate::error::Result<()> {
        self.run("DrawObjCreatorMultiCurve")
    }

    /// `DrawObjCreatorMultiLine` — 다각선 그리기
    pub fn draw_obj_creator_multi_line(&self) -> crate::error::Result<()> {
        self.run("DrawObjCreatorMultiLine")
    }

    /// `DrawObjCreatorPolygon` — 다각형 그리기
    pub fn draw_obj_creator_polygon(&self) -> crate::error::Result<()> {
        self.run("DrawObjCreatorPolygon")
    }

    /// `DrawObjCreatorStar` — 별 그리기
    pub fn draw_obj_creator_star(&self) -> crate::error::Result<()> {
        self.run("DrawObjCreatorStar")
    }

    /// `DrawObjCreatorFreeDrawing` — 자유 그리기
    pub fn draw_obj_creator_free_drawing(&self) -> crate::error::Result<()> {
        self.run("DrawObjCreatorFreeDrawing")
    }

    /// `DrawObjCreatorCanvas` — 묶음 개체(캔버스) 그리기
    pub fn draw_obj_creator_canvas(&self) -> crate::error::Result<()> {
        self.run("DrawObjCreatorCanvas")
    }

    /// `DrawObjCreatorTextBox` — 글상자 그리기
    pub fn draw_obj_creator_text_box(&self) -> crate::error::Result<()> {
        self.run("DrawObjCreatorTextBox")
    }

    /// `DrawObjCreatorHorzTextBox` — 가로 글상자 그리기
    pub fn draw_obj_creator_horz_text_box(&self) -> crate::error::Result<()> {
        self.run("DrawObjCreatorHorzTextBox")
    }

    /// `DrawObjCreatorPicture` — 그림 그리기
    pub fn draw_obj_creator_picture(&self) -> crate::error::Result<()> {
        self.run("DrawObjCreatorPicture")
    }

    /// `DrawObjCreatorOleObject` — OLE 개체 그리기
    pub fn draw_obj_creator_ole_object(&self) -> crate::error::Result<()> {
        self.run("DrawObjCreatorOleObject")
    }

    /// `DrawObjCreatorVideo` — 비디오 개체 그리기
    pub fn draw_obj_creator_video(&self) -> crate::error::Result<()> {
        self.run("DrawObjCreatorVideo")
    }

    // ── 모양 복사 ──

    /// `ShapeCopyPaste` — 모양 복사 (ParameterSet: `ShapeCopyPaste`)
    pub fn shape_copy_paste(&self) -> crate::error::Result<()> {
        self.run("ShapeCopyPaste")
    }

    /// `ShapeObjectCopy` — 그리기 모양 복사 (ParameterSet: `ShapeObjectCopyPaste`)
    pub fn shape_object_copy(&self) -> crate::error::Result<()> {
        self.run("ShapeObjectCopy")
    }

    /// `ShapeObjectPaste` — 그리기 모양 붙여넣기 (ParameterSet: `ShapeObjectCopyPaste`)
    pub fn shape_object_paste(&self) -> crate::error::Result<()> {
        self.run("ShapeObjectPaste")
    }

    // ── 개체 선택·도구 ──

    /// `ShapeObjSelect` — 틀 선택 도구
    pub fn shape_obj_select(&self) -> crate::error::Result<()> {
        self.run("ShapeObjSelect")
    }

    /// `ShapeObjNextObject` — 이후 개체로 이동 (tab 키)
    pub fn shape_obj_next_object(&self) -> crate::error::Result<()> {
        self.run("ShapeObjNextObject")
    }

    /// `ShapeObjPrevObject` — 이전 개체로 이동 (shift + tab 키)
    pub fn shape_obj_prev_object(&self) -> crate::error::Result<()> {
        self.run("ShapeObjPrevObject")
    }

    /// `ShapeObjTableSelCell` — 테이블 선택상태에서 첫 번째 셀 선택하기
    pub fn shape_obj_table_sel_cell(&self) -> crate::error::Result<()> {
        self.run("ShapeObjTableSelCell")
    }

    // ── 속성 대화상자 ──

    /// `ShapeObjAttrDialog` — 틀 속성 환경설정 (ParameterSet: `ShapeObject`)
    pub fn shape_obj_attr_dialog(&self) -> crate::error::Result<()> {
        self.run("ShapeObjAttrDialog")
    }

    /// `ShapeObjDialog` — 환경설정 (ParameterSet: `ShapeObject`)
    pub fn shape_obj_dialog(&self) -> crate::error::Result<()> {
        self.run("ShapeObjDialog")
    }

    /// `ShapeObjComment` — 개체 설명문 수정하기 (ParameterSet: `ShapeObject`)
    pub fn shape_obj_comment(&self) -> crate::error::Result<()> {
        self.run("ShapeObjComment")
    }

    // ── 순서 ──

    /// `ShapeObjBringToFront` — 맨 앞으로
    pub fn shape_obj_bring_to_front(&self) -> crate::error::Result<()> {
        self.run("ShapeObjBringToFront")
    }

    /// `ShapeObjBringForward` — 앞으로
    pub fn shape_obj_bring_forward(&self) -> crate::error::Result<()> {
        self.run("ShapeObjBringForward")
    }

    /// `ShapeObjBringInFrontOfText` — 글 앞으로
    pub fn shape_obj_bring_in_front_of_text(&self) -> crate::error::Result<()> {
        self.run("ShapeObjBringInFrontOfText")
    }

    /// `ShapeObjSendToBack` — 맨 뒤로
    pub fn shape_obj_send_to_back(&self) -> crate::error::Result<()> {
        self.run("ShapeObjSendToBack")
    }

    /// `ShapeObjSendBack` — 뒤로
    pub fn shape_obj_send_back(&self) -> crate::error::Result<()> {
        self.run("ShapeObjSendBack")
    }

    /// `ShapeObjCtrlSendBehindText` — 글 뒤로
    pub fn shape_obj_ctrl_send_behind_text(&self) -> crate::error::Result<()> {
        self.run("ShapeObjCtrlSendBehindText")
    }

    // ── 정렬 ──

    /// `ShapeObjAlignLeft` — 왼쪽으로 정렬
    pub fn shape_obj_align_left(&self) -> crate::error::Result<()> {
        self.run("ShapeObjAlignLeft")
    }

    /// `ShapeObjAlignCenter` — 가운데로 정렬
    pub fn shape_obj_align_center(&self) -> crate::error::Result<()> {
        self.run("ShapeObjAlignCenter")
    }

    /// `ShapeObjAlignRight` — 오른쪽으로 정렬
    pub fn shape_obj_align_right(&self) -> crate::error::Result<()> {
        self.run("ShapeObjAlignRight")
    }

    /// `ShapeObjAlignTop` — 위로 정렬
    pub fn shape_obj_align_top(&self) -> crate::error::Result<()> {
        self.run("ShapeObjAlignTop")
    }

    /// `ShapeObjAlignMiddle` — 중간 정렬
    pub fn shape_obj_align_middle(&self) -> crate::error::Result<()> {
        self.run("ShapeObjAlignMiddle")
    }

    /// `ShapeObjAlignBottom` — 아래로 정렬
    pub fn shape_obj_align_bottom(&self) -> crate::error::Result<()> {
        self.run("ShapeObjAlignBottom")
    }

    /// `ShapeObjAlignWidth` — 폭 맞춤
    pub fn shape_obj_align_width(&self) -> crate::error::Result<()> {
        self.run("ShapeObjAlignWidth")
    }

    /// `ShapeObjAlignHeight` — 높이 맞춤
    pub fn shape_obj_align_height(&self) -> crate::error::Result<()> {
        self.run("ShapeObjAlignHeight")
    }

    /// `ShapeObjAlignSize` — 폭/높이 맞춤
    pub fn shape_obj_align_size(&self) -> crate::error::Result<()> {
        self.run("ShapeObjAlignSize")
    }

    /// `ShapeObjAlignHorzSpacing` — 왼쪽/오른쪽 일정한 비율로 정렬
    pub fn shape_obj_align_horz_spacing(&self) -> crate::error::Result<()> {
        self.run("ShapeObjAlignHorzSpacing")
    }

    /// `ShapeObjAlignVertSpacing` — 위/아래 일정한 비율로 정렬
    pub fn shape_obj_align_vert_spacing(&self) -> crate::error::Result<()> {
        self.run("ShapeObjAlignVertSpacing")
    }

    // ── 그룹 ──

    /// `ShapeObjGroup` — 틀 묶기
    pub fn shape_obj_group(&self) -> crate::error::Result<()> {
        self.run("ShapeObjGroup")
    }

    /// `ShapeObjUngroup` — 틀 풀기
    pub fn shape_obj_ungroup(&self) -> crate::error::Result<()> {
        self.run("ShapeObjUngroup")
    }

    // ── 잠금·기본 ──

    /// `ShapeObjLock` — 개체 Lock
    pub fn shape_obj_lock(&self) -> crate::error::Result<()> {
        self.run("ShapeObjLock")
    }

    /// `ShapeObjUnlockAll` — 개체 Unlock All
    pub fn shape_obj_unlock_all(&self) -> crate::error::Result<()> {
        self.run("ShapeObjUnlockAll")
    }

    /// `ShapeObjNorm` — 기본 도형 설정
    pub fn shape_obj_norm(&self) -> crate::error::Result<()> {
        self.run("ShapeObjNorm")
    }

    /// `ShapeObjProtectSize` — 그리기 개체 크기 고정 (ParameterSet: `ShapeObject`)
    pub fn shape_obj_protect_size(&self) -> crate::error::Result<()> {
        self.run("ShapeObjProtectSize")
    }

    // ── 키보드 이동·크기 조절 ──

    /// `ShapeObjMoveLeft` — 키로 움직이기 (왼쪽)
    pub fn shape_obj_move_left(&self) -> crate::error::Result<()> {
        self.run("ShapeObjMoveLeft")
    }

    /// `ShapeObjMoveRight` — 키로 움직이기 (오른쪽)
    pub fn shape_obj_move_right(&self) -> crate::error::Result<()> {
        self.run("ShapeObjMoveRight")
    }

    /// `ShapeObjMoveUp` — 키로 움직이기 (위)
    pub fn shape_obj_move_up(&self) -> crate::error::Result<()> {
        self.run("ShapeObjMoveUp")
    }

    /// `ShapeObjMoveDown` — 키로 움직이기 (아래)
    pub fn shape_obj_move_down(&self) -> crate::error::Result<()> {
        self.run("ShapeObjMoveDown")
    }

    /// `ShapeObjResizeLeft` — 키로 크기 조절 (shift + 왼쪽)
    pub fn shape_obj_resize_left(&self) -> crate::error::Result<()> {
        self.run("ShapeObjResizeLeft")
    }

    /// `ShapeObjResizeRight` — 키로 크기 조절 (shift + 오른쪽)
    pub fn shape_obj_resize_right(&self) -> crate::error::Result<()> {
        self.run("ShapeObjResizeRight")
    }

    /// `ShapeObjResizeUp` — 키로 크기 조절 (shift + 위)
    pub fn shape_obj_resize_up(&self) -> crate::error::Result<()> {
        self.run("ShapeObjResizeUp")
    }

    /// `ShapeObjResizeDown` — 키로 크기 조절 (shift + 아래)
    pub fn shape_obj_resize_down(&self) -> crate::error::Result<()> {
        self.run("ShapeObjResizeDown")
    }

    // ── 회전 ──

    /// `ShapeObjRightAngleRotater` — 90도 회전
    pub fn shape_obj_right_angle_rotater(&self) -> crate::error::Result<()> {
        self.run("ShapeObjRightAngleRotater")
    }

    /// `ShapeObjRightAngleRotaterAnticlockwise` — -90도 회전
    pub fn shape_obj_right_angle_rotater_anticlockwise(&self) -> crate::error::Result<()> {
        self.run("ShapeObjRightAngleRotaterAnticlockwise")
    }

    /// `ShapeObjRandomAngleRotater` — 자유각 회전 (ParameterSet: `ShapeObject`)
    pub fn shape_obj_random_angle_rotater(&self) -> crate::error::Result<()> {
        self.run("ShapeObjRandomAngleRotater")
    }

    /// `ShapeObjRotater` — 자유각 회전 (회전중심 고정)
    pub fn shape_obj_rotater(&self) -> crate::error::Result<()> {
        self.run("ShapeObjRotater")
    }

    // ── 뒤집기 ──

    /// `ShapeObjHorzFlip` — 그리기 개체 좌우 뒤집기
    pub fn shape_obj_horz_flip(&self) -> crate::error::Result<()> {
        self.run("ShapeObjHorzFlip")
    }

    /// `ShapeObjHorzFlipOrgState` — 그리기 개체 좌우 뒤집기 원상태로 되돌리기
    pub fn shape_obj_horz_flip_org_state(&self) -> crate::error::Result<()> {
        self.run("ShapeObjHorzFlipOrgState")
    }

    /// `ShapeObjVertFlip` — 그리기 개체 상하 뒤집기
    pub fn shape_obj_vert_flip(&self) -> crate::error::Result<()> {
        self.run("ShapeObjVertFlip")
    }

    /// `ShapeObjVertFlipOrgState` — 그리기 개체 상하 뒤집기 원상태로 되돌리기
    pub fn shape_obj_vert_flip_org_state(&self) -> crate::error::Result<()> {
        self.run("ShapeObjVertFlipOrgState")
    }

    // ── 기울이기·음영 ──

    /// `ShapeObjShear` — 그리기 개체 기울이기 (ParameterSet: `ShapeObject`)
    pub fn shape_obj_shear(&self) -> crate::error::Result<()> {
        self.run("ShapeObjShear")
    }

    /// `ShapeObjNoShade` — 채우기 색 음영 없음 (ParameterSet: `ShapeObject`)
    pub fn shape_obj_no_shade(&self) -> crate::error::Result<()> {
        self.run("ShapeObjNoShade")
    }

    /// `ShapeObjNoShadow` — 그림자 없음 (ParameterSet: `ShapeObject`)
    pub fn shape_obj_no_shadow(&self) -> crate::error::Result<()> {
        self.run("ShapeObjNoShadow")
    }

    // ── 텍스트박스 ──

    /// `ShapeObjAttachTextBox` — 글 상자로 만들기
    pub fn shape_obj_attach_text_box(&self) -> crate::error::Result<()> {
        self.run("ShapeObjAttachTextBox")
    }

    /// `ShapeObjDetachTextBox` — 글상자 속성 없애기
    pub fn shape_obj_detach_text_box(&self) -> crate::error::Result<()> {
        self.run("ShapeObjDetachTextBox")
    }

    /// `ShapeObjTextBoxEdit` — 글상자 선택상태에서 편집모드로 들어가기
    pub fn shape_obj_text_box_edit(&self) -> crate::error::Result<()> {
        self.run("ShapeObjTextBoxEdit")
    }

    /// `ShapeObjToggleTextBox` — 도형 글 상자로 만들기 Toggle
    pub fn shape_obj_toggle_text_box(&self) -> crate::error::Result<()> {
        self.run("ShapeObjToggleTextBox")
    }

    // ── 캡션 ──

    /// `ShapeObjAttachCaption` — 캡션 넣기
    pub fn shape_obj_attach_caption(&self) -> crate::error::Result<()> {
        self.run("ShapeObjAttachCaption")
    }

    /// `ShapeObjDetachCaption` — 캡션 없애기
    pub fn shape_obj_detach_caption(&self) -> crate::error::Result<()> {
        self.run("ShapeObjDetachCaption")
    }

    /// `ShapeObjInsertCaptionNum` — 캡션 번호 넣기
    pub fn shape_obj_insert_caption_num(&self) -> crate::error::Result<()> {
        self.run("ShapeObjInsertCaptionNum")
    }

    // ── 텍스트 감싸기 ──

    /// `ShapeObjWrapSquare` — 직사각형 텍스트 감싸기
    pub fn shape_obj_wrap_square(&self) -> crate::error::Result<()> {
        self.run("ShapeObjWrapSquare")
    }

    /// `ShapeObjWrapTopAndBottom` — 자리 차지 텍스트 감싸기
    pub fn shape_obj_wrap_top_and_bottom(&self) -> crate::error::Result<()> {
        self.run("ShapeObjWrapTopAndBottom")
    }

    // ── 안내선 ──

    /// `ShapeObjGuideLine` — 그리기 개체 안내선
    pub fn shape_obj_guide_line(&self) -> crate::error::Result<()> {
        self.run("ShapeObjGuideLine")
    }

    /// `ShapeObjShowGuideLine` — 그리기 개체 안내선 표시
    pub fn shape_obj_show_guide_line(&self) -> crate::error::Result<()> {
        self.run("ShapeObjShowGuideLine")
    }

    /// `ShapeObjShowGuideLineBase` — 그리기 안내선 (한글 2024 이상)
    pub fn shape_obj_show_guide_line_base(&self) -> crate::error::Result<()> {
        self.run("ShapeObjShowGuideLineBase")
    }

    // ── 선 속성 ──

    /// `ShapeObjFillProperty` — 고치기 대화상자중 fill tab
    pub fn shape_obj_fill_property(&self) -> crate::error::Result<()> {
        self.run("ShapeObjFillProperty")
    }

    /// `ShapeObjLineProperty` — 고치기 대화상자중 line tab
    pub fn shape_obj_line_property(&self) -> crate::error::Result<()> {
        self.run("ShapeObjLineProperty")
    }

    /// `ShapeObjLineStyleOhter` — 다른 선 종류
    pub fn shape_obj_line_style_other(&self) -> crate::error::Result<()> {
        self.run("ShapeObjLineStyleOhter")
    }

    /// `ShapeObjLineWidthOhter` — 다른 선 굵기
    pub fn shape_obj_line_width_other(&self) -> crate::error::Result<()> {
        self.run("ShapeObjLineWidthOhter")
    }

    // ── 그림으로 저장 ──

    /// `ShapeObjSaveAsPicture` — 그리기개체를 그림으로 저장하기
    pub fn shape_obj_save_as_picture(&self) -> crate::error::Result<()> {
        self.run("ShapeObjSaveAsPicture")
    }

    // ── 글맵시(TextArt) ──

    /// `TextArtShadowMobeToDown` — 글맵시 그림자 위치 이동-아래로 (ParameterSet: `ShapeObject`)
    pub fn text_art_shadow_move_to_down(&self) -> crate::error::Result<()> {
        self.run("TextArtShadowMobeToDown")
    }

    /// `TextArtShadowMobeToLeft` — 글맵시 그림자 위치 이동-왼쪽으로 (ParameterSet: `ShapeObject`)
    pub fn text_art_shadow_move_to_left(&self) -> crate::error::Result<()> {
        self.run("TextArtShadowMobeToLeft")
    }

    /// `TextArtShadowMobeToRight` — 글맵시 그림자 위치 이동-오른쪽으로 (ParameterSet: `ShapeObject`)
    pub fn text_art_shadow_move_to_right(&self) -> crate::error::Result<()> {
        self.run("TextArtShadowMobeToRight")
    }

    /// `TextArtShadowMoveToUp` — 글맵시 그림자 위치 이동-위로 (ParameterSet: `ShapeObject`)
    pub fn text_art_shadow_move_to_up(&self) -> crate::error::Result<()> {
        self.run("TextArtShadowMoveToUp")
    }

    // ── 글상자 정렬 ──

    /// `TextBoxAlignLeftTop` — 글상자 정렬 (ParameterSet: `ShapeObject`)
    pub fn text_box_align_left_top(&self) -> crate::error::Result<()> {
        self.run("TextBoxAlignLeftTop")
    }

    /// `TextBoxAlignLeftCenter` — 글상자 정렬 (ParameterSet: `ShapeObject`)
    pub fn text_box_align_left_center(&self) -> crate::error::Result<()> {
        self.run("TextBoxAlignLeftCenter")
    }

    /// `TextBoxAlignLeftBottom` — 글상자 정렬 (ParameterSet: `ShapeObject`)
    pub fn text_box_align_left_bottom(&self) -> crate::error::Result<()> {
        self.run("TextBoxAlignLeftBottom")
    }

    /// `TextBoxAlignCenterTop` — 글상자 정렬 (ParameterSet: `ShapeObject`)
    pub fn text_box_align_center_top(&self) -> crate::error::Result<()> {
        self.run("TextBoxAlignCenterTop")
    }

    /// `TextBoxAlignCenterCenter` — 글상자 정렬 (ParameterSet: `ShapeObject`)
    pub fn text_box_align_center_center(&self) -> crate::error::Result<()> {
        self.run("TextBoxAlignCenterCenter")
    }

    /// `TextBoxAlignCenterBottom` — 글상자 정렬 (ParameterSet: `ShapeObject`)
    pub fn text_box_align_center_bottom(&self) -> crate::error::Result<()> {
        self.run("TextBoxAlignCenterBottom")
    }

    /// `TextBoxAlignRightTop` — 글상자 정렬 (ParameterSet: `ShapeObject`)
    pub fn text_box_align_right_top(&self) -> crate::error::Result<()> {
        self.run("TextBoxAlignRightTop")
    }

    /// `TextBoxAlignRightCenter` — 글상자 정렬 (ParameterSet: `ShapeObject`)
    pub fn text_box_align_right_center(&self) -> crate::error::Result<()> {
        self.run("TextBoxAlignRightCenter")
    }

    /// `TextBoxAlignRightBottom` — 글상자 정렬 (ParameterSet: `ShapeObject`)
    pub fn text_box_align_right_bottom(&self) -> crate::error::Result<()> {
        self.run("TextBoxAlignRightBottom")
    }

    // ── 글상자 세로 정렬 ──

    /// `TextBoxVAlignTop` — 글상자 세로 정렬-위 (ParameterSet: `ShapeObject`)
    pub fn text_box_v_align_top(&self) -> crate::error::Result<()> {
        self.run("TextBoxVAlignTop")
    }

    /// `TextBoxVAlignCenter` — 글상자 세로 정렬-가운데 (ParameterSet: `ShapeObject`)
    pub fn text_box_v_align_center(&self) -> crate::error::Result<()> {
        self.run("TextBoxVAlignCenter")
    }

    /// `TextBoxVAlignBottom` — 글상자 세로 정렬-아래 (ParameterSet: `ShapeObject`)
    pub fn text_box_v_align_bottom(&self) -> crate::error::Result<()> {
        self.run("TextBoxVAlignBottom")
    }

    // ── 글상자 문자 방향 ──

    /// `TextBoxTextHorz` — 글상자 문자 방향-가로 쓰기 (ParameterSet: `ShapeObject`)
    pub fn text_box_text_horz(&self) -> crate::error::Result<()> {
        self.run("TextBoxTextHorz")
    }

    /// `TextBoxTextVert` — 글상자 문자 방향-세로 쓰기-영문 눕힘 (ParameterSet: `ShapeObject`)
    pub fn text_box_text_vert(&self) -> crate::error::Result<()> {
        self.run("TextBoxTextVert")
    }

    /// `TextBoxTextVertAll` — 글상자 문자 방향-세로 쓰기-영문 세움 (ParameterSet: `ShapeObject`)
    pub fn text_box_text_vert_all(&self) -> crate::error::Result<()> {
        self.run("TextBoxTextVertAll")
    }

    /// `TextBoxToggleDirection` — 글상자 문자 방향-세로/가로 토글 (ParameterSet: `ShapeObject`)
    pub fn text_box_toggle_direction(&self) -> crate::error::Result<()> {
        self.run("TextBoxToggleDirection")
    }
}
