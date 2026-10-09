/// 커서 이동 관련 액션
///
/// SDK 참고: ActionTable_2504.pdf § Move*
use super::Act;

impl Act<'_> {
    /// `MoveWordBegin` — 커서를 현재 단어의 시작으로 이동합니다.
    pub fn move_word_begin(&self) -> crate::error::Result<()> {
        self.0.h_action()?.run("MoveWordBegin")
    }
}
