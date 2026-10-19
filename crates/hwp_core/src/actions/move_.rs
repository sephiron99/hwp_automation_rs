use crate::method::h_action::HAction;

/// 커서 이동 관련 액션
///
/// SDK 참고: ActionTable_2504.pdf § Move*
impl HAction {
    /// `MoveWordBegin` — 커서를 현재 단어의 시작으로 이동합니다.
    pub fn move_word_begin(&self) -> crate::error::Result<()> {
        self.run("MoveWordBegin")
    }
}
