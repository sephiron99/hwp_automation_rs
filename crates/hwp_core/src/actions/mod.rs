/// Action 층 — ActionTable의 Action ID별 Rust 래퍼
///
/// SDK 참고: ActionTable_2504.pdf
///
/// 모든 래퍼는 [`HwpObject::act`]가 돌려주는 [`Act`] 핸들의 메서드다.
///
/// ```ignore
/// hwp.act().undo()?;
/// ```
///
/// # 구조
/// - [`app`] — 앱·파일·인쇄·프레젠테이션·버전
/// - [`edit`] — 편집·선택·클립보드·실행취소
/// - [`find`] — 찾기·바꾸기·검색
/// - [`chars`] — 글자 모양
/// - [`para`] — 문단 모양·스타일·글머리표
/// - [`move_`] — 커서 이동 (일반/선택 포함)
/// - [`insert`] — 삽입 (그림·자동번호·필드 등)
/// - [`table`] — 표 조작
/// - [`draw`] — 그리기 개체·도형
/// - [`picture`] — 그림 속성·효과
/// - [`view`] — 보기 옵션·화면 배율
/// - [`page`] — 편집 용지·바탕쪽·구역
/// - [`note`] — 각주·미주·주석
/// - [`track`] — 변경 추적
/// - [`macro_`] — 매크로·빠른 교정·빠른 책갈피
pub mod app;
pub mod chars;
pub mod draw;
pub mod edit;
pub mod find;
pub mod insert;
pub mod macro_;
pub mod move_;
pub mod note;
pub mod page;
pub mod para;
pub mod picture;
pub mod table;
pub mod track;
pub mod view;

use crate::hwp_obj::HwpObject;

/// Action 래퍼의 진입점. [`HwpObject::act`]로 얻는다.
pub struct Act<'a>(&'a HwpObject);

impl Act<'_> {
    /// Action ID를 이름으로 직접 실행합니다. 전용 래퍼가 없는 Action에 사용합니다.
    pub fn run(&self, action_id: &str) -> crate::error::Result<()> {
        self.0.h_action()?.run(action_id)
    }
}

impl HwpObject {
    /// Action 래퍼 핸들을 반환합니다.
    pub fn act(&self) -> Act<'_> {
        Act(self)
    }
}
