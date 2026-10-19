//! Action 층 — ActionTable의 Action ID별 Rust 래퍼
//!
//! SDK 참고: ActionTable_2504.pdf
//!
//! 모든 래퍼는 [`HAction`](crate::method::h_action::HAction)의 메서드다. 이 모듈의 하위 파일들이
//! 도메인별로 `impl HAction` 블록을 나눠 정의한다. `HAction`은 [`HwpObject::h_action`]으로 얻는다.
//! 전용 래퍼가 없는 Action은 [`HAction::run`](crate::method::h_action::HAction::run)으로 실행한다.
//!
//! ```ignore
//! hwp.h_action()?.undo()?;
//!
//! // 여러 Action을 실행할 때는 HAction을 한 번만 얻는다.
//! let action = hwp.h_action()?;
//! action.select_all()?;
//! action.copy()?;
//! action.run("FileNew")?; // 래퍼가 없는 Action
//! ```
//!
//! [`HwpObject::h_action`]: crate::HwpObject::h_action
//!
//! # 구조
//! - `app` — 앱·파일·인쇄·프레젠테이션·버전
//! - `edit` — 편집·선택·클립보드·실행취소
//! - `find` — 찾기·바꾸기·검색
//! - `chars` — 글자 모양
//! - `para` — 문단 모양·스타일·글머리표
//! - `move_` — 커서 이동 (일반/선택 포함)
//! - `insert` — 삽입 (그림·자동번호·필드 등)
//! - `table` — 표 조작
//! - `draw` — 그리기 개체·도형
//! - `picture` — 그림 속성·효과
//! - `view` — 보기 옵션·화면 배율
//! - `page` — 편집 용지·바탕쪽·구역
//! - `note` — 각주·미주·주석
//! - `track` — 변경 추적
//! - `macro_` — 매크로·빠른 교정·빠른 책갈피
mod app;
mod chars;
mod draw;
mod edit;
mod find;
mod insert;
mod macro_;
mod move_;
mod note;
mod page;
mod para;
mod picture;
mod table;
mod track;
mod view;
