/// 편집·선택·클립보드·실행취소 관련 액션
///
/// SDK 참고: ActionTable_2504.pdf § Cancel, Copy, Cut, Delete*, Paste*, Select*, Undo, Redo 등
use crate::hwp_obj::HwpObject;

impl HwpObject {
    // ── 취소·실행취소·다시 실행 ──

    /// `Cancel` — 현재 작업을 취소합니다.
    pub fn cancel(&self) -> crate::error::Result<()> {
        self.h_action()?.run("Cancel")
    }

    /// `Undo` — 되살리기 (실행 취소)
    pub fn undo(&self) -> crate::error::Result<()> {
        self.h_action()?.run("Undo")
    }

    /// `Redo` — 다시 실행
    pub fn redo(&self) -> crate::error::Result<()> {
        self.h_action()?.run("Redo")
    }

    // ── 클립보드 ──

    /// `Copy` — 복사합니다.
    pub fn copy(&self) -> crate::error::Result<()> {
        self.h_action()?.run("Copy")
    }

    /// `Cut` — 잘라냅니다.
    pub fn cut(&self) -> crate::error::Result<()> {
        self.h_action()?.run("Cut")
    }

    /// `Paste` — 붙이기
    pub fn paste(&self) -> crate::error::Result<()> {
        self.h_action()?.run("Paste")
    }

    /// `PastePage` — 쪽 붙여넣기
    pub fn paste_page(&self) -> crate::error::Result<()> {
        self.h_action()?.run("PastePage")
    }

    /// `PasteSpecial` — 골라 붙이기 대화상자를 표시합니다.
    pub fn paste_special(&self) -> crate::error::Result<()> {
        self.h_action()?.run("PasteSpecial")
    }

    /// `ClipboardHistoryCopy` — 클립보드 히스토리 복사
    pub fn clipboard_history_copy(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ClipboardHistoryCopy")
    }

    /// `ClipboardHistoryDlg` — 클립보드 히스토리 대화상자를 표시합니다.
    pub fn clipboard_history_dlg(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ClipboardHistoryDlg")
    }

    /// `ClipboardHistoryPaste` — 클립보드 히스토리 붙이기
    pub fn clipboard_history_paste(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ClipboardHistoryPaste")
    }

    // ── 삭제 ──

    /// `Delete` — Delete 키와 동일합니다.
    pub fn delete(&self) -> crate::error::Result<()> {
        self.h_action()?.run("Delete")
    }

    /// `DeleteBack` — Backspace 키와 동일합니다.
    pub fn delete_back(&self) -> crate::error::Result<()> {
        self.h_action()?.run("DeleteBack")
    }

    /// `DeleteLine` — 현재 줄을 삭제합니다.
    pub fn delete_line(&self) -> crate::error::Result<()> {
        self.h_action()?.run("DeleteLine")
    }

    /// `DeleteLineEnd` — 커서부터 줄 끝까지 삭제합니다.
    pub fn delete_line_end(&self) -> crate::error::Result<()> {
        self.h_action()?.run("DeleteLineEnd")
    }

    /// `DeleteWord` — 다음 단어를 삭제합니다.
    pub fn delete_word(&self) -> crate::error::Result<()> {
        self.h_action()?.run("DeleteWord")
    }

    /// `DeleteWordBack` — 이전 단어를 삭제합니다.
    pub fn delete_word_back(&self) -> crate::error::Result<()> {
        self.h_action()?.run("DeleteWordBack")
    }

    /// `Erase` — 선택 영역을 지웁니다.
    pub fn erase(&self) -> crate::error::Result<()> {
        self.h_action()?.run("Erase")
    }

    // ── 선택 ──

    /// `SelectAll` — 현재 리스트에서 모두 선택
    pub fn select_all(&self) -> crate::error::Result<()> {
        self.h_action()?.run("SelectAll")
    }

    /// `SelectColumn` — 칸 블록 선택 (F4 키와 동일)
    pub fn select_column(&self) -> crate::error::Result<()> {
        self.h_action()?.run("SelectColumn")
    }

    /// `SelectCtrlFront` — 개체 선택 정방향
    pub fn select_ctrl_front(&self) -> crate::error::Result<()> {
        self.h_action()?.run("SelectCtrlFront")
    }

    /// `SelectCtrlReverse` — 개체 선택 역방향
    pub fn select_ctrl_reverse(&self) -> crate::error::Result<()> {
        self.h_action()?.run("SelectCtrlReverse")
    }

    // ── 블록 이동 ──

    /// `LeftShiftBlock` — 텍스트 블록을 왼쪽으로 이동합니다.
    pub fn left_shift_block(&self) -> crate::error::Result<()> {
        self.h_action()?.run("LeftShiftBlock")
    }

    /// `RightShiftBlock` — 텍스트 블록을 오른쪽으로 이동합니다.
    pub fn right_shift_block(&self) -> crate::error::Result<()> {
        self.h_action()?.run("RightShiftBlock")
    }

    // ── 기타 편집 ──

    /// `ToggleOverwrite` — 수정/삽입 모드를 전환합니다.
    pub fn toggle_overwrite(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ToggleOverwrite")
    }

    /// `ReturnKeyInField` — 필드 안에서 Return 키에 대한 액션
    pub fn return_key_in_field(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ReturnKeyInField")
    }

    /// `ReturnPrevPos` — 직전 위치로 돌아가기
    pub fn return_prev_pos(&self) -> crate::error::Result<()> {
        self.h_action()?.run("ReturnPrevPos")
    }

    /// `RecalcPageCount` — 현재 페이지의 쪽 번호를 재계산합니다.
    pub fn recalc_page_count(&self) -> crate::error::Result<()> {
        self.h_action()?.run("RecalcPageCount")
    }

    /// `NextTextBoxLinked` — 연결된 글상자의 다음 글상자로 이동합니다.
    pub fn next_text_box_linked(&self) -> crate::error::Result<()> {
        self.h_action()?.run("NextTextBoxLinked")
    }

    /// `PrevTextBoxLinked` — 연결된 글상자의 이전 글상자로 이동합니다.
    pub fn prev_text_box_linked(&self) -> crate::error::Result<()> {
        self.h_action()?.run("PrevTextBoxLinked")
    }

    /// `UnlinkTextBox` — 글상자 연결을 끊습니다.
    pub fn unlink_text_box(&self) -> crate::error::Result<()> {
        self.h_action()?.run("UnlinkTextBox")
    }

    // ── 정렬·계산 ──

    /// `Sort` — 소트 대화상자를 표시합니다.
    pub fn sort(&self) -> crate::error::Result<()> {
        self.h_action()?.run("Sort")
    }

    /// `Sum` — 블록 합계를 계산합니다.
    pub fn sum(&self) -> crate::error::Result<()> {
        self.h_action()?.run("Sum")
    }
}
