//! 문서 전체 텍스트(컨트롤 포함) 추출용 lending-free iterator.
//!
//! `InitScan` / `GetText` / `ReleaseScan` 시퀀스를 표준 [`Iterator`]와 RAII
//! guard로 감싼다. 호출자는 [`HwpTextExt::text_segments`]로 [`TextWalker`]를
//! 얻어 `for seg in walker { ... }` 형태로 순회하면 된다. walker가 Drop될
//! 때 자동으로 `ReleaseScan`이 호출된다.
//!
//! 본문뿐 아니라 머리말/꼬리말/각주/표/텍스트박스 등 컨트롤 내부 텍스트도
//! 포함되며, 컨트롤 진입/탈출은 [`TextSegment::EnterCtrl`] / [`TextSegment::ExitCtrl`]
//! 이벤트로 노출된다.
//!
//! # 사용 예
//! ```ignore
//! use std::ops::ControlFlow;
//! use hwp_addon::text_extract::{HwpTextExt, TextSegment};
//!
//! // 구조 정보가 필요한 경우 (문단/컨트롤 경계 구분)
//! for seg in hwp.text_segments()? {
//!     match seg? {
//!         TextSegment::Text(s) => print!("{s}"),
//!         TextSegment::ParaBreak => println!(),
//!         TextSegment::EnterCtrl(name) => println!("[enter {name}]"),
//!         TextSegment::ExitCtrl => println!("[exit]"),
//!     }
//! }
//!
//! // 전체를 단일 String으로 (구조 정보 무시)
//! let full = hwp.get_all_text()?;
//! ```

use std::ops::ControlFlow;

use hwp_core::error::{HwpError, Result};
use hwp_core::hwp_obj::HwpObject;
use hwp_core::ihwpobject::lib::{GetTextStatus, ScanEpos, ScanRange, ScanSpos, mask};

/// 한 텍스트 추출 이벤트.
#[derive(Debug, Clone)]
pub enum TextSegment {
    /// 본문/컨트롤 내부 텍스트 청크.
    Text(String),
    /// 문단 경계.
    ParaBreak,
    /// 컨트롤 진입. SDK(`HwpAutomation_2504.pdf` § GetText, status=4)는
    /// 이때 반환되는 BSTR의 내용을 명시하지 않는다 — 빈 문자열일 수 있다.
    EnterCtrl(String),
    /// 컨트롤 종료.
    ExitCtrl,
}

/// `Iterator::next()`의 상태 머신.
///
/// 한 번의 `get_text` 호출이 두 이벤트(예: `Text` + `ParaBreak`)를 만들 수
/// 있으므로, 보류된 이벤트와 종료 상태를 따로 추적한다. 각 변형 문서 참고.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Pending {
    /// 다음 `next()`는 `get_text`를 호출해야 한다.
    None,
    /// 직전에 `NextParagraph`였고 텍스트는 이미 yield 했으므로,
    /// 다음 `next()`는 `get_text` 없이 `ParaBreak`만 yield한다.
    ParaBreakDeferred,
    /// 스캔 종료. 모든 후속 호출은 `None` 반환.
    Done,
}

/// 문서 텍스트 walker. Drop 시 자동으로 `ReleaseScan`을 호출한다.
pub struct TextWalker<'h> {
    hwp: &'h HwpObject,
    pending: Pending,
    finished: bool,
}

impl<'h> TextWalker<'h> {
    fn new(
        hwp: &'h HwpObject,
        option: u32,
        range: ScanRange,
        spara: u32,
        spos: u32,
        epara: u32,
        epos: u32,
    ) -> Result<Self> {
        let ok = hwp.init_scan(option, range, spara, spos, epara, epos)?;
        if !ok {
            return Err(HwpError::ExecutionFailed("InitScan returned false".into()));
        }
        Ok(Self {
            hwp,
            pending: Pending::None,
            finished: false,
        })
    }

    fn release(&mut self) {
        if !self.finished {
            self.finished = true;
            if let Err(e) = self.hwp.release_scan() {
                hwp_core::debug::log("text_extract", &format!("release_scan failed: {e}"));
            }
        }
    }
}

impl Iterator for TextWalker<'_> {
    type Item = Result<TextSegment>;

    fn next(&mut self) -> Option<Result<TextSegment>> {
        match self.pending {
            Pending::Done => return None,
            Pending::ParaBreakDeferred => {
                self.pending = Pending::None;
                return Some(Ok(TextSegment::ParaBreak));
            }
            Pending::None => {}
        }

        let (status, text) = match self.hwp.get_text() {
            Ok(v) => v,
            Err(e) => {
                self.pending = Pending::Done;
                self.release();
                return Some(Err(e));
            }
        };

        match status {
            GetTextStatus::Normal => Some(Ok(TextSegment::Text(text))),
            GetTextStatus::NextParagraph => {
                if text.is_empty() {
                    Some(Ok(TextSegment::ParaBreak))
                } else {
                    self.pending = Pending::ParaBreakDeferred;
                    Some(Ok(TextSegment::Text(text)))
                }
            }
            GetTextStatus::EnterControl => Some(Ok(TextSegment::EnterCtrl(text))),
            GetTextStatus::ExitControl => Some(Ok(TextSegment::ExitCtrl)),
            GetTextStatus::EndOfList | GetTextStatus::None => {
                self.pending = Pending::Done;
                self.release();
                None
            }
            GetTextStatus::NotInitialized => {
                self.pending = Pending::Done;
                self.release();
                Some(Err(HwpError::ExecutionFailed(
                    "GetText: not initialized".into(),
                )))
            }
            GetTextStatus::ConversionFailed => {
                self.pending = Pending::Done;
                self.release();
                Some(Err(HwpError::ExecutionFailed(
                    "GetText: conversion failed".into(),
                )))
            }
            GetTextStatus::Unknown(n) => {
                self.pending = Pending::Done;
                self.release();
                Some(Err(HwpError::ExecutionFailed(format!(
                    "GetText: unknown status {n}"
                ))))
            }
        }
    }
}

impl Drop for TextWalker<'_> {
    fn drop(&mut self) {
        self.release();
    }
}

/// [`HwpObject`]에 컨트롤 포함 문서 전체 텍스트 추출 메서드를 추가하는 extension trait.
pub trait HwpTextExt {
    /// 문서 전체에 대한 walker. 기본 옵션:
    /// - mask = `NORMAL | CHAR | INLINE | CTRL` (모든 컨트롤 포함)
    /// - range = `Document → Document`
    fn text_segments(&self) -> Result<TextWalker<'_>>;

    /// 마스크/범위를 직접 지정해 walker를 만든다.
    ///
    /// # 매개변수
    /// - `option` — 스캔 대상 마스크. [`mask`] 상수의 OR 조합
    ///   (예: `mask::NORMAL | mask::CTRL`).
    /// - `range` — 검색 범위. [`ScanRange::new`]로 시작/끝 위치 종류를
    ///   지정하거나 [`ScanRange::within_selection`]으로 선택 블록에 한정한다.
    /// - `spara`, `spos` — 시작 문단 번호와 그 안에서의 문자 위치.
    ///   `range`의 시작이 [`ScanSpos::Specified`]일 때만 의미가 있고,
    ///   그 외에는 0을 넘기면 된다.
    /// - `epara`, `epos` — 끝 문단/문자 위치. `range`의 끝이
    ///   [`ScanEpos::Specified`]일 때만 의미가 있고, 그 외에는 0.
    ///
    /// 내부적으로 [`HwpObject::init_scan`]에 그대로 전달된다.
    fn text_segments_with(
        &self,
        option: u32,
        range: ScanRange,
        spara: u32,
        spos: u32,
        epara: u32,
        epos: u32,
    ) -> Result<TextWalker<'_>>;

    /// 전체 텍스트를 단일 `String`으로 반환한다. 컨트롤 진입/탈출 경계는
    /// 무시되고 컨트롤 내부 텍스트도 본문과 이어붙는다. 문단 경계는 `'\n'`.
    fn get_all_text(&self) -> Result<String>;

    /// closure 기반 visitor. `ControlFlow::Break(())`를 돌려주면 즉시 종료.
    fn visit_all_text<F>(&self, f: F) -> Result<()>
    where
        F: FnMut(&TextSegment) -> ControlFlow<()>;
}

impl HwpTextExt for HwpObject {
    fn text_segments(&self) -> Result<TextWalker<'_>> {
        self.text_segments_with(
            mask::NORMAL | mask::CHAR | mask::INLINE | mask::CTRL,
            ScanRange::new(ScanSpos::Document, ScanEpos::Document),
            0,
            0,
            0,
            0,
        )
    }

    fn text_segments_with(
        &self,
        option: u32,
        range: ScanRange,
        spara: u32,
        spos: u32,
        epara: u32,
        epos: u32,
    ) -> Result<TextWalker<'_>> {
        TextWalker::new(self, option, range, spara, spos, epara, epos)
    }

    fn get_all_text(&self) -> Result<String> {
        let mut out = String::new();
        for seg in self.text_segments()? {
            match seg? {
                TextSegment::Text(s) | TextSegment::EnterCtrl(s) => out.push_str(&s),
                TextSegment::ParaBreak => out.push('\n'),
                TextSegment::ExitCtrl => {}
            }
        }
        Ok(out)
    }

    fn visit_all_text<F>(&self, mut f: F) -> Result<()>
    where
        F: FnMut(&TextSegment) -> ControlFlow<()>,
    {
        for seg in self.text_segments()? {
            let seg = seg?;
            if let ControlFlow::Break(()) = f(&seg) {
                break;
            }
        }
        Ok(())
    }
}
