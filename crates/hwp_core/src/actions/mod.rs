/// ActionTable 전체에 대한 Rust 바인딩
///
/// SDK 참고: ActionTable_2504.pdf
///
/// # 구조
/// - [`app`] — 앱·파일·인쇄·프레젠테이션·버전
/// - [`edit`] — 편집·선택·클립보드·실행취소
/// - [`find`] — 찾기·바꾸기·검색
/// - [`char`] — 글자 모양
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
pub mod note;
pub mod page;
pub mod para;
pub mod picture;
pub mod table;
pub mod track;
pub mod view;
