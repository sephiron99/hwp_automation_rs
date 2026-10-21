//! HWPML2X → LLM이 읽기 좋은 간결한 텍스트
//!
//! `GetTextFile("HWPML2X")`이 돌려주는 XML은 글자·문단 모양을 `HEAD`의 아이디로만 가리키고,
//! 화면 배치용 요소와 기본값 속성이 많아 LLM에 그대로 넘기기에는 크다. 이 모듈은 HWPML을
//! 한 번 훑어 다음 형식으로 바꾼다.
//!
//! ```text
//! [글자 모양]
//! c20 = 한양신명조 10pt
//! c27 = 함초롬바탕 9pt 굵게 색 #C00000
//! [문단 모양]
//! p1 = 양쪽 정렬
//! [본문]
//! ¶p1 {c20}⟨표1⟩제1절 수사의 의의와 ⟦CARET⟧조건
//!   ⟨표1⟩ 1행×1열
//!     [1,1]
//!       ¶p1 {c25}케이스1
//! ```
//!
//! - 모양 목록에는 출력한 본문에서 실제로 쓴 모양만, 사람이 읽는 값으로 적는다.
//! - 본문 한 줄이 문단 하나다. `{cN}`은 그 자리부터 글자 모양이 `cN`이라는 뜻이다.
//! - 표·머리말·각주처럼 문단 안의 개체는 그 자리에 `⟨표1⟩` 같은 표시를 두고, 내용은 문단
//!   줄 아래에 들여 써서 적는다.
//! - 커서 책갈피는 `⟦CARET⟧`로 바꾼다.
//!
//! 한글을 부르지 않는 순수 변환이라 HWPML 문자열만 있으면 어디서든 쓸 수 있다. 이 모듈은
//! 표준 라이브러리와 `quick-xml`만 쓴다.
//!
//! 근거: 한글문서파일형식3.0_HWPML_revision1.2.pdf(4.3 글자·문단 모양, 5 본문 엘리먼트)
use std::collections::{BTreeSet, HashMap};
use std::fmt::Write as _;

use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader, XmlVersion};

/// 커서 자리를 나타내는 표식
pub const CARET_MARK: &str = "⟦CARET⟧";

/// [`hwpml_to_llm_text`]의 옵션
#[derive(Debug, Clone, Default)]
pub struct LlmTextOptions {
    /// 이 이름의 책갈피를 [`CARET_MARK`]로 바꾼다. `None`이면 커서를 표시하지 않는다.
    pub caret_bookmark: Option<String>,
    /// `Some(n)`이면 커서가 든 본문 문단의 앞뒤 `n`개 문단만 내보내고 나머지는 생략 줄로
    /// 줄인다. 커서를 찾지 못하면 전체를 내보낸다.
    pub context_paragraphs: Option<usize>,
}

/// HWPML을 해석하지 못했을 때의 에러
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct HwpmlParseError(pub String);

impl std::fmt::Display for HwpmlParseError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        f.write_str(&self.0)
    }
}

impl std::error::Error for HwpmlParseError {}

type Result<T> = std::result::Result<T, HwpmlParseError>;

/// HWPML2X 문자열을 LLM용 텍스트로 바꾼다. 형식은 [모듈 설명](self) 참고.
pub fn hwpml_to_llm_text(hwpml: &str, options: &LlmTextOptions) -> Result<String> {
    let mut reader = Reader::from_str(hwpml);
    let mut head = Head::default();
    let mut body: Vec<Para> = Vec::new();
    let caret = options.caret_bookmark.as_deref();

    loop {
        match reader.read_event().map_err(xml_err)? {
            Event::Start(e) if e.name().as_ref() == "HEAD" => head = parse_head(&mut reader)?,
            Event::Start(e) if e.name().as_ref() == "SECTION" => {
                body.extend(parse_paras_until(&mut reader, "SECTION", caret)?);
            }
            Event::Eof => break,
            _ => {}
        }
    }

    Ok(render(&head, &body, options))
}

// =========================================================================
// HEAD: 글꼴, 글자 모양, 문단 모양, 스타일
// =========================================================================

#[derive(Default)]
struct Head {
    /// 언어별(`Hangul`, `Latin` ..) 글꼴 아이디 → 이름
    fonts: HashMap<String, HashMap<u32, String>>,
    char_shapes: HashMap<u32, CharShape>,
    para_shapes: HashMap<u32, ParaShape>,
    styles: HashMap<u32, String>,
}

#[derive(Default)]
struct CharShape {
    height: i64,
    text_color: u32,
    shade_color: u32,
    font_hangul: Option<u32>,
    font_latin: Option<u32>,
    ratio: i64,
    spacing: i64,
    bold: bool,
    italic: bool,
    underline: bool,
    strikeout: bool,
    superscript: bool,
    subscript: bool,
    outline: bool,
    shadow: bool,
    emboss: bool,
    engrave: bool,
}

#[derive(Default)]
struct ParaShape {
    align: String,
    heading_type: String,
    level: u32,
    indent: String,
    left: String,
    right: String,
    prev: String,
    next: String,
    line_spacing_type: String,
    line_spacing: String,
}

fn parse_head(reader: &mut Reader<&[u8]>) -> Result<Head> {
    let mut head = Head::default();
    let mut lang = String::new();
    let mut cur_char: Option<(u32, CharShape)> = None;
    let mut cur_para: Option<(u32, ParaShape)> = None;

    loop {
        let event = reader.read_event().map_err(xml_err)?;
        let (e, is_start) = match &event {
            Event::Start(e) => (e, true),
            Event::Empty(e) => (e, false),
            Event::End(e) => {
                match e.name().as_ref() {
                    "HEAD" => break,
                    "CHARSHAPE" => {
                        if let Some((id, cs)) = cur_char.take() {
                            head.char_shapes.insert(id, cs);
                        }
                    }
                    "PARASHAPE" => {
                        if let Some((id, ps)) = cur_para.take() {
                            head.para_shapes.insert(id, ps);
                        }
                    }
                    _ => {}
                }
                continue;
            }
            Event::Eof => return Err(HwpmlParseError("HEAD가 닫히지 않았습니다".into())),
            _ => continue,
        };
        match e.name().as_ref() {
            "FONTFACE" => lang = attr(e, "Lang")?.unwrap_or_default(),
            "FONT" => {
                if let (Some(id), Some(name)) = (num::<u32>(e, "Id")?, attr(e, "Name")?) {
                    head.fonts.entry(lang.clone()).or_default().insert(id, name);
                }
            }
            "CHARSHAPE" => {
                let cs = CharShape {
                    height: num(e, "Height")?.unwrap_or(1000),
                    text_color: num(e, "TextColor")?.unwrap_or(0),
                    shade_color: num(e, "ShadeColor")?.unwrap_or(u32::MAX),
                    ratio: 100,
                    ..CharShape::default()
                };
                let id = num(e, "Id")?.unwrap_or(0);
                if is_start {
                    cur_char = Some((id, cs));
                } else {
                    head.char_shapes.insert(id, cs);
                }
            }
            name if cur_char.is_some() => {
                let cs = &mut cur_char.as_mut().expect("is_some").1;
                match name {
                    "FONTID" => {
                        cs.font_hangul = num(e, "Hangul")?;
                        cs.font_latin = num(e, "Latin")?;
                    }
                    "RATIO" => cs.ratio = num(e, "Hangul")?.unwrap_or(100),
                    "CHARSPACING" => cs.spacing = num(e, "Hangul")?.unwrap_or(0),
                    "BOLD" => cs.bold = true,
                    "ITALIC" => cs.italic = true,
                    "UNDERLINE" => cs.underline = attr(e, "Type")?.as_deref() != Some("None"),
                    "STRIKEOUT" => cs.strikeout = attr(e, "Type")?.as_deref() != Some("None"),
                    "SUPERSCRIPT" => cs.superscript = true,
                    "SUBSCRIPT" => cs.subscript = true,
                    "OUTLINE" => cs.outline = attr(e, "Type")?.as_deref() != Some("None"),
                    "SHADOW" => cs.shadow = attr(e, "Type")?.as_deref() != Some("None"),
                    "EMBOSS" => cs.emboss = true,
                    "ENGRAVE" => cs.engrave = true,
                    _ => {}
                }
            }
            "PARASHAPE" => {
                let ps = ParaShape {
                    align: attr(e, "Align")?.unwrap_or_else(|| "Justify".into()),
                    heading_type: attr(e, "HeadingType")?.unwrap_or_else(|| "None".into()),
                    level: num(e, "Level")?.unwrap_or(0),
                    ..ParaShape::default()
                };
                let id = num(e, "Id")?.unwrap_or(0);
                if is_start {
                    cur_para = Some((id, ps));
                } else {
                    head.para_shapes.insert(id, ps);
                }
            }
            "PARAMARGIN" => {
                if let Some((_, ps)) = cur_para.as_mut() {
                    ps.indent = attr(e, "Indent")?.unwrap_or_default();
                    ps.left = attr(e, "Left")?.unwrap_or_default();
                    ps.right = attr(e, "Right")?.unwrap_or_default();
                    ps.prev = attr(e, "Prev")?.unwrap_or_default();
                    ps.next = attr(e, "Next")?.unwrap_or_default();
                    ps.line_spacing_type =
                        attr(e, "LineSpacingType")?.unwrap_or_else(|| "Percent".into());
                    ps.line_spacing = attr(e, "LineSpacing")?.unwrap_or_else(|| "160".into());
                }
            }
            "STYLE" => {
                if let (Some(id), Some(name)) = (num::<u32>(e, "Id")?, attr(e, "Name")?) {
                    head.styles.insert(id, name);
                }
            }
            _ => {}
        }
    }
    Ok(head)
}

// =========================================================================
// BODY: 문단, 문단 안의 개체
// =========================================================================

struct Para {
    para_shape: u32,
    style: u32,
    inlines: Vec<Inline>,
    blocks: Vec<Block>,
    has_caret: bool,
}

enum Inline {
    Text {
        char_shape: u32,
        text: String,
    },
    Caret,
    /// 개체 자리 표시. `blocks`의 같은 번호 블록과 짝을 이룬다(내용이 없는 개체는 블록 없음).
    Object(ObjectRef),
}

#[derive(Clone, Copy)]
struct ObjectRef {
    kind: ObjectKind,
    /// 같은 종류 안에서의 번호. 렌더링 때 문서 순서대로 다시 매긴다.
    block: Option<usize>,
}

#[derive(Clone, Copy, PartialEq, Eq, Hash)]
enum ObjectKind {
    Table,
    Header,
    Footer,
    Footnote,
    Endnote,
    Picture,
    Equation,
    Drawing,
    Ole,
    Field,
    Other,
}

impl ObjectKind {
    fn from_element(name: &str) -> Option<Self> {
        Some(match name {
            "TABLE" => Self::Table,
            "HEADER" => Self::Header,
            "FOOTER" => Self::Footer,
            "FOOTNOTE" => Self::Footnote,
            "ENDNOTE" => Self::Endnote,
            "PICTURE" => Self::Picture,
            "EQUATION" => Self::Equation,
            "OLE" => Self::Ole,
            "CONTAINER" | "TEXTART" | "LINE" | "RECTANGLE" | "ELLIPSE" | "ARC" | "POLYGON"
            | "CURVE" | "CONNECTLINE" => Self::Drawing,
            "BUTTON" | "RADIOBUTTON" | "CHECKBUTTON" | "COMBOBOX" | "EDIT" | "LISTBOX"
            | "SCROLLBAR" => Self::Field,
            // 화면 배치나 번호처럼 글 내용이 아닌 컨트롤
            "SECDEF" | "COLDEF" | "FIELDBEGIN" | "FIELDEND" | "BOOKMARK" | "AUTONUM" | "NEWNUM"
            | "PAGENUMCTRL" | "PAGEHIDING" | "PAGENUM" | "INDEXMARK" => return None,
            _ => Self::Other,
        })
    }

    fn label(self) -> &'static str {
        match self {
            Self::Table => "표",
            Self::Header => "머리말",
            Self::Footer => "꼬리말",
            Self::Footnote => "각주",
            Self::Endnote => "미주",
            Self::Picture => "그림",
            Self::Equation => "수식",
            Self::Drawing => "그리기 개체",
            Self::Ole => "OLE 개체",
            Self::Field => "양식 개체",
            Self::Other => "개체",
        }
    }
}

enum Block {
    Table {
        rows: u32,
        cols: u32,
        cells: Vec<Cell>,
    },
    /// 머리말·각주·글상자처럼 문단 목록을 품은 개체
    Lists(Vec<Vec<Para>>),
}

struct Cell {
    row: u32,
    col: u32,
    row_span: u32,
    col_span: u32,
    paras: Vec<Para>,
}

/// `end` 요소가 닫힐 때까지 그 안의 `P`를 읽는다.
fn parse_paras_until(
    reader: &mut Reader<&[u8]>,
    end: &str,
    caret: Option<&str>,
) -> Result<Vec<Para>> {
    let mut paras = Vec::new();
    loop {
        match reader.read_event().map_err(xml_err)? {
            Event::Start(e) if e.name().as_ref() == "P" => {
                paras.push(parse_para(reader, &e, caret)?);
            }
            Event::Empty(e) if e.name().as_ref() == "P" => paras.push(empty_para(&e)?),
            Event::End(e) if e.name().as_ref() == end => return Ok(paras),
            Event::Eof => return Err(unclosed(end)),
            _ => {}
        }
    }
}

fn empty_para(e: &BytesStart<'_>) -> Result<Para> {
    Ok(Para {
        para_shape: num(e, "ParaShape")?.unwrap_or(0),
        style: num(e, "Style")?.unwrap_or(0),
        inlines: Vec::new(),
        blocks: Vec::new(),
        has_caret: false,
    })
}

fn parse_para(
    reader: &mut Reader<&[u8]>,
    start: &BytesStart<'_>,
    caret: Option<&str>,
) -> Result<Para> {
    let mut para = empty_para(start)?;
    let mut char_shape = 0u32;
    // CHAR 안에 있는 동안의 글
    let mut in_char = false;

    loop {
        let event = reader.read_event().map_err(xml_err)?;
        match event {
            Event::End(e) => match e.name().as_ref() {
                "P" => break,
                "CHAR" => in_char = false,
                _ => {}
            },
            Event::Eof => return Err(unclosed("P")),
            Event::Text(t) if in_char => push_text(&mut para, char_shape, &t.xml10_content()),
            Event::GeneralRef(r) if in_char => {
                let s = match r.resolve_char_ref().map_err(xml_err)? {
                    Some(c) => c.to_string(),
                    None => quick_xml::escape::resolve_predefined_entity(&r)
                        .unwrap_or("")
                        .to_owned(),
                };
                push_text(&mut para, char_shape, &s);
            }
            Event::Start(e) | Event::Empty(e) if matches!(e.name().as_ref(), "TEXT") => {
                char_shape = num(&e, "CharShape")?.unwrap_or(0);
            }
            Event::Start(e) if e.name().as_ref() == "CHAR" => in_char = true,
            Event::Empty(e) if e.name().as_ref() == "CHAR" => {}
            Event::Start(e) | Event::Empty(e) if in_char => {
                // CHAR 안의 특수 문자
                let s = match e.name().as_ref() {
                    "TAB" => "\t",
                    "LINEBREAK" => "↵",
                    "HYPEN" => "-",
                    "NBSPACE" => "\u{00A0}",
                    "FWSPACE" => "\u{3000}",
                    _ => "",
                };
                push_text(&mut para, char_shape, s);
            }
            Event::Empty(e) if e.name().as_ref() == "BOOKMARK" => {
                if caret.is_some() && attr(&e, "Name")?.as_deref() == caret {
                    para.inlines.push(Inline::Caret);
                    para.has_caret = true;
                }
            }
            Event::Start(e) => {
                let name = e.name().as_ref().to_owned();
                let kind = ObjectKind::from_element(&name);
                let block = if name == "TABLE" {
                    Some(parse_table(reader, &e, caret)?)
                } else {
                    parse_object(reader, &name, caret)?
                };
                let has_caret = block.as_ref().is_some_and(block_has_caret);
                para.has_caret |= has_caret;
                // 내용 없는 배치용 컨트롤은 커서가 들어 있지 않으면 표시하지 않는다.
                let kind = match (kind, &block) {
                    (Some(k), _) => k,
                    (None, Some(_)) => ObjectKind::Other,
                    (None, None) => continue,
                };
                let block = block.map(|b| {
                    para.blocks.push(b);
                    para.blocks.len() - 1
                });
                para.inlines.push(Inline::Object(ObjectRef { kind, block }));
            }
            Event::Empty(e) => {
                if let Some(kind) = ObjectKind::from_element(e.name().as_ref()) {
                    para.inlines
                        .push(Inline::Object(ObjectRef { kind, block: None }));
                }
            }
            _ => {}
        }
    }
    Ok(para)
}

fn push_text(para: &mut Para, char_shape: u32, s: &str) {
    if s.is_empty() {
        return;
    }
    if let Some(Inline::Text {
        char_shape: cs,
        text,
    }) = para.inlines.last_mut()
        && *cs == char_shape
    {
        text.push_str(s);
        return;
    }
    para.inlines.push(Inline::Text {
        char_shape,
        text: s.to_owned(),
    });
}

fn parse_table(
    reader: &mut Reader<&[u8]>,
    start: &BytesStart<'_>,
    caret: Option<&str>,
) -> Result<Block> {
    let rows = num(start, "RowCount")?.unwrap_or(0);
    let cols = num(start, "ColCount")?.unwrap_or(0);
    let mut cells = Vec::new();
    let mut cell: Option<Cell> = None;
    loop {
        match reader.read_event().map_err(xml_err)? {
            Event::Start(e) if e.name().as_ref() == "CELL" => {
                cell = Some(Cell {
                    row: num(&e, "RowAddr")?.unwrap_or(0),
                    col: num(&e, "ColAddr")?.unwrap_or(0),
                    row_span: num(&e, "RowSpan")?.unwrap_or(1),
                    col_span: num(&e, "ColSpan")?.unwrap_or(1),
                    paras: Vec::new(),
                });
            }
            Event::Start(e) if e.name().as_ref() == "PARALIST" => {
                let paras = parse_paras_until(reader, "PARALIST", caret)?;
                if let Some(c) = cell.as_mut() {
                    c.paras.extend(paras);
                }
            }
            Event::End(e) => match e.name().as_ref() {
                "CELL" => cells.extend(cell.take()),
                "TABLE" => break,
                _ => {}
            },
            Event::Eof => return Err(unclosed("TABLE")),
            _ => {}
        }
    }
    Ok(Block::Table { rows, cols, cells })
}

/// 개체 하나를 끝까지 읽고, 그 안의 문단 목록(`PARALIST`)을 돌려준다. 없으면 `None`.
fn parse_object(
    reader: &mut Reader<&[u8]>,
    name: &str,
    caret: Option<&str>,
) -> Result<Option<Block>> {
    let mut lists = Vec::new();
    let mut depth = 0usize;
    loop {
        match reader.read_event().map_err(xml_err)? {
            Event::Start(e) if e.name().as_ref() == "PARALIST" => {
                lists.push(parse_paras_until(reader, "PARALIST", caret)?);
            }
            Event::Start(e) if e.name().as_ref() == "TABLE" => {
                // 글상자 안의 표처럼 문단 밖에 바로 놓인 표는 드물지만 문단으로 감싸 보존한다.
                let table = parse_table(reader, &e, caret)?;
                lists.push(vec![Para {
                    para_shape: 0,
                    style: 0,
                    has_caret: block_has_caret(&table),
                    inlines: vec![Inline::Object(ObjectRef {
                        kind: ObjectKind::Table,
                        block: Some(0),
                    })],
                    blocks: vec![table],
                }]);
            }
            Event::Start(_) => depth += 1,
            Event::End(e) => {
                if depth == 0 && e.name().as_ref() == name {
                    break;
                }
                depth = depth.saturating_sub(1);
            }
            Event::Eof => return Err(unclosed(name)),
            _ => {}
        }
    }
    Ok((!lists.is_empty()).then_some(Block::Lists(lists)))
}

fn block_has_caret(block: &Block) -> bool {
    match block {
        Block::Table { cells, .. } => cells.iter().any(|c| c.paras.iter().any(|p| p.has_caret)),
        Block::Lists(lists) => lists.iter().flatten().any(|p| p.has_caret),
    }
}

// =========================================================================
// 렌더링
// =========================================================================

#[derive(Default)]
struct Renderer {
    out: String,
    used_chars: BTreeSet<u32>,
    used_paras: BTreeSet<u32>,
    counters: HashMap<ObjectKind, usize>,
    /// 설명이 같은 모양은 가장 작은 아이디 하나로 합친다(번호 정의만 다른 문단 모양 등).
    char_alias: HashMap<u32, u32>,
    para_alias: HashMap<u32, u32>,
}

/// 아이디 → 설명이 같은 모양 가운데 가장 작은 아이디
fn alias_by_description(
    ids: impl Iterator<Item = u32>,
    describe: impl Fn(u32) -> String,
) -> HashMap<u32, u32> {
    let mut ids: Vec<u32> = ids.collect();
    ids.sort_unstable();
    let mut first: HashMap<String, u32> = HashMap::new();
    ids.into_iter()
        .map(|id| (id, *first.entry(describe(id)).or_insert(id)))
        .collect()
}

fn render(head: &Head, body: &[Para], options: &LlmTextOptions) -> String {
    let caret_at = body.iter().position(|p| p.has_caret);
    let range = match (options.context_paragraphs, caret_at) {
        (Some(n), Some(i)) => i.saturating_sub(n)..(i + n + 1).min(body.len()),
        _ => 0..body.len(),
    };

    let mut r = Renderer {
        char_alias: alias_by_description(head.char_shapes.keys().copied(), |id| {
            describe_char(head, id)
        }),
        para_alias: alias_by_description(head.para_shapes.keys().copied(), |id| {
            describe_para(head, id)
        }),
        ..Renderer::default()
    };
    if range.start > 0 {
        let _ = writeln!(r.out, "… (앞 문단 {}개 생략)", range.start);
    }
    for para in &body[range.clone()] {
        r.para(head, para, 0);
    }
    if range.end < body.len() {
        let _ = writeln!(r.out, "… (뒤 문단 {}개 생략)", body.len() - range.end);
    }
    if options.caret_bookmark.is_some() && caret_at.is_none() {
        let _ = writeln!(r.out, "(커서 위치를 찾지 못했습니다)");
    }

    let mut doc = String::with_capacity(r.out.len() + 1024);
    doc.push_str("[글자 모양]\n");
    for id in &r.used_chars {
        let _ = writeln!(doc, "c{id} = {}", describe_char(head, *id));
    }
    doc.push_str("[문단 모양]\n");
    for id in &r.used_paras {
        let _ = writeln!(doc, "p{id} = {}", describe_para(head, *id));
    }
    doc.push_str("[본문]\n");
    doc.push_str(&r.out);
    doc
}

impl Renderer {
    fn para(&mut self, head: &Head, para: &Para, depth: usize) {
        let indent = "  ".repeat(depth);
        let para_shape = *self
            .para_alias
            .get(&para.para_shape)
            .unwrap_or(&para.para_shape);
        self.used_paras.insert(para_shape);
        let _ = write!(self.out, "{indent}¶p{para_shape}");
        if para.style != 0
            && let Some(name) = head.styles.get(&para.style)
        {
            let _ = write!(self.out, " «{name}»");
        }
        if !para.inlines.is_empty() {
            self.out.push(' ');
        }

        // 개체 번호는 문서 순서대로 매긴다. 블록을 적을 때 같은 번호를 쓴다.
        let mut block_labels: Vec<(usize, String)> = Vec::new();
        let mut last_char: Option<u32> = None;
        for inline in &para.inlines {
            match inline {
                Inline::Text { char_shape, text } => {
                    let char_shape = *self.char_alias.get(char_shape).unwrap_or(char_shape);
                    if last_char != Some(char_shape) {
                        self.used_chars.insert(char_shape);
                        let _ = write!(self.out, "{{c{char_shape}}}");
                        last_char = Some(char_shape);
                    }
                    self.out.push_str(text);
                }
                Inline::Caret => self.out.push_str(CARET_MARK),
                Inline::Object(obj) => {
                    let n = self.counters.entry(obj.kind).or_insert(0);
                    *n += 1;
                    let label = format!("{}{}", obj.kind.label(), n);
                    let _ = write!(self.out, "⟨{label}⟩");
                    if let Some(b) = obj.block {
                        block_labels.push((b, label));
                    }
                }
            }
        }
        self.out.push('\n');

        for (b, label) in block_labels {
            self.block(head, &para.blocks[b], &label, depth + 1);
        }
    }

    fn block(&mut self, head: &Head, block: &Block, label: &str, depth: usize) {
        let indent = "  ".repeat(depth);
        match block {
            Block::Table { rows, cols, cells } => {
                let _ = writeln!(self.out, "{indent}⟨{label}⟩ {rows}행×{cols}열");
                for cell in cells {
                    let _ = write!(self.out, "{indent}  [{},{}]", cell.row + 1, cell.col + 1);
                    if cell.row_span > 1 || cell.col_span > 1 {
                        let _ = write!(self.out, " 병합 {}행×{}열", cell.row_span, cell.col_span);
                    }
                    self.out.push('\n');
                    for p in &cell.paras {
                        self.para(head, p, depth + 2);
                    }
                }
            }
            Block::Lists(lists) => {
                let _ = writeln!(self.out, "{indent}⟨{label}⟩");
                for p in lists.iter().flatten() {
                    self.para(head, p, depth + 1);
                }
            }
        }
    }
}

fn describe_char(head: &Head, id: u32) -> String {
    let Some(cs) = head.char_shapes.get(&id) else {
        return "(정의 없음)".into();
    };
    let font =
        |lang: &str, fid: Option<u32>| fid.and_then(|f| head.fonts.get(lang)?.get(&f).cloned());
    let hangul = font("Hangul", cs.font_hangul);
    let latin = font("Latin", cs.font_latin);
    let mut s = match (&hangul, &latin) {
        (Some(h), Some(l)) if h != l => format!("{h}/{l}"),
        (Some(h), _) => h.clone(),
        (None, Some(l)) => l.clone(),
        (None, None) => "글꼴?".into(),
    };
    let _ = write!(s, " {}pt", hundredths(cs.height));
    for (on, word) in [
        (cs.bold, "굵게"),
        (cs.italic, "기울임"),
        (cs.underline, "밑줄"),
        (cs.strikeout, "취소선"),
        (cs.superscript, "위첨자"),
        (cs.subscript, "아래첨자"),
        (cs.outline, "외곽선"),
        (cs.shadow, "그림자"),
        (cs.emboss, "양각"),
        (cs.engrave, "음각"),
    ] {
        if on {
            s.push(' ');
            s.push_str(word);
        }
    }
    if cs.ratio != 100 {
        let _ = write!(s, " 장평 {}%", cs.ratio);
    }
    if cs.spacing != 0 {
        let _ = write!(s, " 자간 {}%", cs.spacing);
    }
    if cs.text_color != 0 {
        let _ = write!(s, " 색 {}", rgb(cs.text_color));
    }
    if cs.shade_color != u32::MAX && cs.shade_color != 0xFF_FFFF {
        let _ = write!(s, " 음영 {}", rgb(cs.shade_color));
    }
    s
}

fn describe_para(head: &Head, id: u32) -> String {
    let Some(ps) = head.para_shapes.get(&id) else {
        return "(정의 없음)".into();
    };
    let mut parts: Vec<String> = Vec::new();
    parts.push(
        match ps.align.as_str() {
            "Justify" => "양쪽 정렬",
            "Left" => "왼쪽 정렬",
            "Right" => "오른쪽 정렬",
            "Center" => "가운데 정렬",
            "Distribute" => "배분 정렬",
            "DistributeSpace" => "나눔 정렬",
            other => other,
        }
        .to_owned(),
    );
    match ps.heading_type.as_str() {
        "Outline" => parts.push(format!("개요 {}수준", ps.level + 1)),
        "Number" => parts.push(format!("번호 문단 {}수준", ps.level + 1)),
        "Bullet" => parts.push("글머리표".into()),
        _ => {}
    }
    for (value, name) in [
        (&ps.left, "왼쪽 여백"),
        (&ps.right, "오른쪽 여백"),
        (&ps.prev, "문단 위"),
        (&ps.next, "문단 아래"),
    ] {
        if let Some(v) = length(value) {
            parts.push(format!("{name} {v}"));
        }
    }
    if let Some(v) = length(ps.indent.trim_start_matches('-')) {
        let word = if ps.indent.starts_with('-') {
            "내어쓰기"
        } else {
            "들여쓰기"
        };
        parts.push(format!("{word} {v}"));
    }
    match ps.line_spacing_type.as_str() {
        "Percent" if ps.line_spacing != "160" && !ps.line_spacing.is_empty() => {
            parts.push(format!("줄간격 {}%", ps.line_spacing));
        }
        "Fixed" => parts.extend(length(&ps.line_spacing).map(|v| format!("줄간격 고정 {v}"))),
        "BetweenLines" => {
            parts.extend(length(&ps.line_spacing).map(|v| format!("줄간격 여백 {v}")));
        }
        "AtLeast" => parts.extend(length(&ps.line_spacing).map(|v| format!("줄간격 최소 {v}"))),
        _ => {}
    }
    parts.join(", ")
}

/// HWPML 길이 값(`hwpunit` 또는 `"Nch"`)을 사람이 읽는 값으로. 0이거나 비면 `None`.
fn length(v: &str) -> Option<String> {
    let v = v.trim();
    if let Some(n) = v.strip_suffix("ch") {
        return (n.trim() != "0" && !n.is_empty()).then(|| format!("{}글자", n.trim()));
    }
    match v.parse::<i64>() {
        Ok(0) | Err(_) => None,
        Ok(n) => Some(format!("{}pt", hundredths(n))),
    }
}

/// 1/100 단위 정수를 `10`, `9.5`, `-2.25`처럼 쓴다.
fn hundredths(n: i64) -> String {
    let s = format!("{:.2}", n as f64 / 100.0);
    s.trim_end_matches('0').trim_end_matches('.').to_owned()
}

/// HWPML 색(`0x00bbggrr`)을 `#RRGGBB`로
fn rgb(c: u32) -> String {
    let (r, g, b) = (c & 0xFF, (c >> 8) & 0xFF, (c >> 16) & 0xFF);
    format!("#{r:02X}{g:02X}{b:02X}")
}

// =========================================================================
// 속성 읽기
// =========================================================================

fn attr(e: &BytesStart<'_>, name: &str) -> Result<Option<String>> {
    for a in e.attributes() {
        let a = a.map_err(|err| HwpmlParseError(err.to_string()))?;
        if a.key.as_ref() == name {
            let v = a
                .normalized_value(XmlVersion::Implicit1_0)
                .map_err(xml_err)?;
            return Ok(Some(v.into_owned()));
        }
    }
    Ok(None)
}

fn num<T: std::str::FromStr>(e: &BytesStart<'_>, name: &str) -> Result<Option<T>> {
    match attr(e, name)? {
        None => Ok(None),
        Some(v) => v
            .trim()
            .parse()
            .map(Some)
            .map_err(|_| HwpmlParseError(format!("{name} 값 '{v}'가 숫자가 아닙니다"))),
    }
}

fn xml_err(e: impl std::fmt::Display) -> HwpmlParseError {
    HwpmlParseError(e.to_string())
}

fn unclosed(name: &str) -> HwpmlParseError {
    HwpmlParseError(format!("{name}가 닫히지 않았습니다"))
}
