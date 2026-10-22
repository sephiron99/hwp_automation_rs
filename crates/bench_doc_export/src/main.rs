//! 현재 문서의 내용과 구조 전체를 얻는 방법별 시간 측정
//!
//! 한글 OLE 클라이언트로 문서를 열고, 같은 문서를 형식별로 `GetTextFile`(메모리 문자열)과
//! `SaveAs`(파일 저장 완료까지)로 내보내는 데 걸린 시간을 잰다. 받은 결과를 해석하는 시간은
//! 재지 않는다.
//!
//! 구조가 빠지는 형식(TEXT, UNICODE, HTML 등)과 바이너리 형식(HWP, MSWORD, PDF)은 대상이
//! 아니다. 남는 것은 HWPML2X다. HWPX는 ZIP으로 묶인 바이너리 파일이지만 구조를 모두 담는
//! 기본 형식이라 `--with-hwpx`를 주면 비교용으로 함께 잰다.
//!
//! ```text
//! bench_doc_export <문서 경로> [--iterations N] [--module 보안모듈이름] [--with-hwpx]
//! ```
//!
//! - 각 경우를 한 번 미리 실행해(첫 호출의 준비 비용 제외) 버린 뒤 N번(기본 5) 잰다.
//! - `--module`을 주면 `RegisterModule("FilePathCheckDLL", 이름)`으로 파일 접근 확인
//!   대화상자를 막는다. 없으면 숨은 한글 창이 대화상자에서 멈출 수 있다.
//! - `SaveAs`는 `lock:false`로 저장해 열린 문서가 저장한 파일로 바뀌지 않게 한다.
//!   저장한 파일은 임시 디렉터리에 두었다가 끝나면 지운다.
//! - `SaveAs+읽기`는 `SaveAs`로 `GetTextFile`과 같은 결과(메모리의 문자열)를 얻기까지의 숨은
//!   비용을 모두 넣어 잰다. 단계별 중앙값도 함께 출력한다.
//!   1. 저장 전 변경 상태 읽기(`SaveAs`가 "변경 안 됨"으로 바꾸므로 되돌리기 위함)
//!   2. `SaveAs(…, "lock:false")`
//!   3. 한글이 쥔 파일 핸들 놓게 하기(`GetTextFile("TEXT")`)
//!   4. 파일을 메모리로 읽고 문자열로 디코딩(BOM이 있으면 UTF-16LE, 없으면 UTF-8)
//!   5. 변경 상태 되돌리기(`set_modified`)
//!   6. 파일 지우기
//! - `SaveAs+읽기 새파일`/`덮어쓰기`는 3단계(핸들 놓게 하기)를 빼고, 한글이 파일을 잡고 있는
//!   채로 읽는다. 새파일은 호출마다 새 경로에 저장하고 앞 호출의 파일을 지운다. 덮어쓰기는
//!   같은 경로에 계속 저장하고 지우지 않는다.
//! - 각 방법이 얻은 문자열이 `GetTextFile` 결과와 같은지, 문서를 고친 뒤 새 내용을 읽는지도
//!   확인한다.
use std::path::{Path, PathBuf};
use std::time::{Duration, Instant};

use hwp_com::HwpObject;
use hwp_com::method::hwp_object::{FileFormat, GetTextFileFormat, MessageBoxMode, SaveOption};

struct Args {
    document: String,
    iterations: usize,
    module: Option<String>,
    with_hwpx: bool,
    /// `SaveAs` 직후 읽기가 왜 느린지 보는 진단. 기다린 시간별 읽기 시간을 잰다.
    probe_dir: Option<String>,
}

fn parse_args() -> Result<Args, String> {
    let mut it = std::env::args().skip(1);
    let mut document = None;
    let mut iterations = 5;
    let mut module = None;
    let mut with_hwpx = false;
    let mut probe_dir = None;
    while let Some(a) = it.next() {
        match a.as_str() {
            "--iterations" => {
                iterations = it
                    .next()
                    .and_then(|v| v.parse().ok())
                    .ok_or("--iterations 뒤에 숫자가 필요합니다")?;
            }
            "--module" => module = Some(it.next().ok_or("--module 뒤에 이름이 필요합니다")?),
            "--with-hwpx" => with_hwpx = true,
            "--probe" => probe_dir = Some(it.next().ok_or("--probe 뒤에 폴더가 필요합니다")?),
            _ if document.is_none() => document = Some(a),
            _ => return Err(format!("알 수 없는 인자: {a}")),
        }
    }
    Ok(Args {
        document: document.ok_or("문서 경로가 필요합니다")?,
        iterations: iterations.max(1),
        module,
        with_hwpx,
        probe_dir,
    })
}

/// 측정할 내보내기 방법
enum Case {
    GetTextFile(GetTextFileFormat),
    SaveAs(FileFormat, &'static str),
    /// `SaveAs`로 메모리의 문자열을 얻기까지 숨은 비용을 모두 넣은 경로
    SaveAsFull(FileFormat, &'static str),
    /// 핸들을 놓게 하지 않고 읽는다. 호출마다 새 파일에 저장하고, 앞 호출의 파일은 이번
    /// `SaveAs`로 핸들이 놓였으므로 그때 지운다.
    SaveAsKeepHandle(FileFormat, &'static str),
    /// 핸들을 놓게 하지 않고 읽는다. 늘 같은 경로에 덮어쓰고 지우지 않는다.
    SaveAsReuse(FileFormat, &'static str),
}

/// 한 번 실행한 결과
struct Measurement {
    total: Duration,
    size: u64,
    /// 얻은 문서 문자열. 내용 비교용(`GetTextFile`과 `SaveAs+읽기` 계열만)
    text: Option<String>,
    /// 단계 이름과 걸린 시간. 단계를 나누지 않는 방법은 비어 있다.
    steps: Vec<(&'static str, Duration)>,
}

impl Case {
    fn name(&self) -> String {
        match self {
            Case::GetTextFile(f) => format!("GetTextFile({})", f.as_str()),
            Case::SaveAs(f, ext) => format!("SaveAs({f}, .{ext})"),
            Case::SaveAsFull(f, ext) => format!("SaveAs+읽기({f}, .{ext})"),
            Case::SaveAsKeepHandle(f, ext) => format!("SaveAs+읽기 새파일({f}, .{ext})"),
            Case::SaveAsReuse(f, ext) => format!("SaveAs+읽기 덮어쓰기({f}, .{ext})"),
        }
    }

    /// 한 번 실행한다. `SaveAs`는 저장한 파일 경로를 `saved`에 남긴다(끝나면 지움).
    fn run(
        &self,
        hwp: &HwpObject,
        tmp: &Path,
        i: usize,
        saved: &mut Vec<PathBuf>,
    ) -> hwp_com::Result<Measurement> {
        match self {
            Case::GetTextFile(f) => {
                let t = Instant::now();
                let s = hwp.get_text_file(*f, false)?;
                Ok(Measurement {
                    total: t.elapsed(),
                    size: s.len() as u64,
                    text: Some(s),
                    steps: Vec::new(),
                })
            }
            Case::SaveAs(f, ext) => {
                let path = tmp.join(format!("bench_{}_{i}.{ext}", f.as_str()));
                let p = path.to_str().expect("임시 경로는 UTF-8");
                let t = Instant::now();
                hwp.save_as(p, f, "lock:false")?;
                let total = t.elapsed();
                saved.push(path.clone());
                Ok(Measurement {
                    total,
                    size: std::fs::metadata(&path).map(|m| m.len()).unwrap_or(0),
                    text: None,
                    steps: Vec::new(),
                })
            }
            Case::SaveAsFull(f, ext) => {
                let path = tmp.join(format!("bench_full_{}_{i}.{ext}", f.as_str()));
                let p = path.to_str().expect("임시 경로는 UTF-8");
                let mut steps = Vec::with_capacity(6);
                let start = Instant::now();
                let mut lap = start;
                let mut mark = |name: &'static str, steps: &mut Vec<_>| {
                    let now = Instant::now();
                    steps.push((name, now - lap));
                    lap = now;
                };

                let doc = hwp.documents()?.active()?;
                let modified = doc.modified()?;
                mark("변경 상태 읽기", &mut steps);
                hwp.save_as(p, f, "lock:false")?;
                mark("SaveAs", &mut steps);
                hwp.get_text_file(GetTextFileFormat::Text, false)?;
                mark("핸들 놓게 하기", &mut steps);
                let bytes = std::fs::read(&path).map_err(io_err)?;
                mark("파일 읽기", &mut steps);
                let text = decode(&bytes);
                mark("디코딩", &mut steps);
                doc.set_modified(modified)?;
                mark("변경 상태 되돌리기", &mut steps);
                std::fs::remove_file(&path).map_err(io_err)?;
                mark("파일 지우기", &mut steps);

                Ok(Measurement {
                    total: start.elapsed(),
                    size: text.len() as u64,
                    text: Some(text),
                    steps,
                })
            }
            Case::SaveAsKeepHandle(f, ext) | Case::SaveAsReuse(f, ext) => {
                let reuse = matches!(self, Case::SaveAsReuse(..));
                let path = if reuse {
                    tmp.join(format!("bench_reuse_{}.{ext}", f.as_str()))
                } else {
                    tmp.join(format!("bench_keep_{}_{i}.{ext}", f.as_str()))
                };
                let p = path.to_str().expect("임시 경로는 UTF-8");
                let mut steps = Vec::with_capacity(6);
                let start = Instant::now();
                let mut lap = start;
                let mut mark = |name: &'static str, steps: &mut Vec<_>| {
                    let now = Instant::now();
                    steps.push((name, now - lap));
                    lap = now;
                };

                let doc = hwp.documents()?.active()?;
                let modified = doc.modified()?;
                mark("변경 상태 읽기", &mut steps);
                hwp.save_as(p, f, "lock:false")?;
                mark("SaveAs", &mut steps);
                let bytes = std::fs::read(&path).map_err(io_err)?;
                mark("파일 읽기(핸들 잡힌 채)", &mut steps);
                let text = decode(&bytes);
                mark("디코딩", &mut steps);
                doc.set_modified(modified)?;
                mark("변경 상태 되돌리기", &mut steps);
                if !reuse {
                    // 앞 호출의 파일은 이번 SaveAs로 핸들이 놓였다.
                    if let Some(prev) = saved.last() {
                        std::fs::remove_file(prev).map_err(io_err)?;
                        saved.pop();
                    }
                    mark("앞 파일 지우기", &mut steps);
                }
                if !saved.contains(&path) {
                    saved.push(path);
                }

                Ok(Measurement {
                    total: start.elapsed(),
                    size: text.len() as u64,
                    text: Some(text),
                    steps,
                })
            }
        }
    }
}

/// 두 문서 문자열이 같은지. 다르면 첫 차이 위치와 앞뒤 일부를 돌려준다.
fn diff(a: &str, b: &str) -> Option<String> {
    let a = a.trim_end_matches('\0');
    let b = b.trim_end_matches('\0');
    if a == b {
        return None;
    }
    let at = a.chars().zip(b.chars()).take_while(|(x, y)| x == y).count();
    let near = |s: &str| {
        s.chars()
            .skip(at.saturating_sub(20))
            .take(60)
            .collect::<String>()
    };
    Some(format!(
        "길이 {} / {}, {at}번째 글자부터 다름: {:?} / {:?}",
        a.chars().count(),
        b.chars().count(),
        near(a),
        near(b)
    ))
}

/// BOM이 있으면 UTF-16LE, 없으면 UTF-8로 디코딩한다.
fn decode(bytes: &[u8]) -> String {
    if let Some(rest) = bytes.strip_prefix(&[0xFF, 0xFE]) {
        let units: Vec<u16> = rest
            .as_chunks::<2>()
            .0
            .iter()
            .map(|c| u16::from_le_bytes(*c))
            .collect();
        String::from_utf16_lossy(&units)
    } else {
        let rest = bytes.strip_prefix(&[0xEF, 0xBB, 0xBF]).unwrap_or(bytes);
        String::from_utf8_lossy(rest).into_owned()
    }
}

fn io_err(e: std::io::Error) -> hwp_com::HwpError {
    hwp_com::HwpError::ExecutionFailed(format!("파일 입출력 실패: {e}"))
}

fn main() {
    let args = match parse_args() {
        Ok(a) => a,
        Err(e) => {
            eprintln!(
                "{e}\n사용법: bench_doc_export <문서 경로> [--iterations N] [--module 이름] [--with-hwpx]"
            );
            std::process::exit(2);
        }
    };
    if let Err(e) = run(&args) {
        eprintln!("실패: {e}");
        std::process::exit(1);
    }
}

fn run(args: &Args) -> hwp_com::Result<()> {
    let hwp = hwp_com::HwpClient::new()?;
    let _ = hwp.set_message_box_mode(
        MessageBoxMode::OK_IDOK
            | MessageBoxMode::OKCANCEL_IDCANCEL
            | MessageBoxMode::YESNO_IDNO
            | MessageBoxMode::YESNOCANCEL_IDCANCEL,
    );
    if let Some(m) = &args.module {
        hwp.register_module("FilePathCheckDLL", m)?;
    }
    let result = bench(&hwp, args);
    let _ = hwp.clear(SaveOption::Discard);
    let _ = hwp.quit();
    result
}

/// `SaveAs` 직후 `wait`만큼 기다린 뒤 파일을 읽는 데 걸린 시간을 잰다.
fn probe(hwp: &HwpObject, dir: &Path, iterations: usize) -> hwp_com::Result<()> {
    for ext in ["hml", "xml"] {
        probe_ext(hwp, dir, iterations, ext)?;
    }
    Ok(())
}

fn probe_ext(hwp: &HwpObject, dir: &Path, iterations: usize, ext: &str) -> hwp_com::Result<()> {
    let ms = |d: Duration| d.as_secs_f64() * 1000.0;
    println!(
        "진단: {} 에 .{ext}로 저장, 기다린 시간별 읽기 시간(중앙값 ms)",
        dir.display()
    );
    let mut saved = Vec::new();
    for wait in [0u64, 5, 10, 20, 40] {
        let mut reads = Vec::new();
        for i in 0..=iterations {
            let path = dir.join(format!("probe_{wait}_{i}.{ext}"));
            hwp.save_as(
                path.to_str().expect("UTF-8"),
                &FileFormat::HWPML2X,
                "lock:false",
            )?;
            std::thread::sleep(Duration::from_millis(wait));
            let t = Instant::now();
            let n = std::fs::read(&path).map_err(io_err)?.len();
            if i > 0 {
                reads.push((t.elapsed(), n));
            }
            saved.push(path);
        }
        reads.sort();
        let (mid, n) = reads[reads.len() / 2];
        println!("  기다림 {wait:>3}ms → 읽기 {:>7.2}ms ({n} bytes)", ms(mid));
    }
    // 비교: 같은 내용을 이 프로세스가 직접 써서 닫은 직후 읽는 시간(백신 검사 등 한글 밖 원인 확인용)
    if let Some(sample) = saved.last() {
        let _ = hwp.get_text_file(GetTextFileFormat::Text, false);
        let bytes = std::fs::read(sample).map_err(io_err)?;
        let mut reads = Vec::new();
        for i in 0..=iterations {
            let path = dir.join(format!("probe_self_{i}.{ext}"));
            std::fs::write(&path, &bytes).map_err(io_err)?;
            let t = Instant::now();
            let _ = std::fs::read(&path).map_err(io_err)?;
            if i > 0 {
                reads.push(t.elapsed());
            }
            saved.push(path);
        }
        reads.sort();
        println!(
            "  직접 쓴 파일, 기다림 0ms → 읽기 {:>7.2}ms",
            ms(reads[reads.len() / 2])
        );
    }
    let _ = hwp.get_text_file(GetTextFileFormat::Text, false);
    for p in saved {
        let _ = std::fs::remove_file(p);
    }
    Ok(())
}

fn bench(hwp: &HwpObject, args: &Args) -> hwp_com::Result<()> {
    hwp.open(&args.document, None, "forceopen:true;suspendpassword:true")?;
    if let Some(dir) = &args.probe_dir {
        return probe(hwp, Path::new(dir), args.iterations);
    }
    let doc_len = std::fs::metadata(&args.document)
        .map(|m| m.len())
        .unwrap_or(0);
    println!(
        "문서: {} ({doc_len} bytes), 반복 {}회",
        args.document, args.iterations
    );

    let mut cases = vec![
        Case::GetTextFile(GetTextFileFormat::HwpMl2X),
        Case::SaveAs(FileFormat::HWPML2X, "hml"),
        Case::SaveAsFull(FileFormat::HWPML2X, "hml"),
        Case::SaveAsKeepHandle(FileFormat::HWPML2X, "hml"),
        Case::SaveAsReuse(FileFormat::HWPML2X, "hml"),
        // 확장자만 .xml로 바꾼 같은 경우들(백신 검사 등 확장자에 따른 차이 확인용)
        Case::SaveAs(FileFormat::HWPML2X, "xml"),
        Case::SaveAsFull(FileFormat::HWPML2X, "xml"),
        Case::SaveAsKeepHandle(FileFormat::HWPML2X, "xml"),
        Case::SaveAsReuse(FileFormat::HWPML2X, "xml"),
    ];
    let reference = hwp.get_text_file(GetTextFileFormat::HwpMl2X, false)?;
    if args.with_hwpx {
        cases.push(Case::SaveAs(FileFormat::HWPX, "hwpx"));
    }

    let tmp = std::env::temp_dir();
    let mut saved = Vec::new();
    let ms = |d: Duration| d.as_secs_f64() * 1000.0;
    println!(
        "{:<24} {:>10} {:>10} {:>10} {:>12}",
        "방법", "최소(ms)", "중앙(ms)", "평균(ms)", "결과(bytes)"
    );
    for case in &cases {
        case.run(hwp, &tmp, 0, &mut saved)?; // 준비 실행
        let mut runs: Vec<Measurement> = Vec::with_capacity(args.iterations);
        for i in 1..=args.iterations {
            runs.push(case.run(hwp, &tmp, i, &mut saved)?);
        }
        let mut times: Vec<Duration> = runs.iter().map(|m| m.total).collect();
        times.sort();
        let mean = times.iter().map(|d| ms(*d)).sum::<f64>() / times.len() as f64;
        println!(
            "{:<24} {:>10.1} {:>10.1} {:>10.1} {:>12}",
            case.name(),
            ms(times[0]),
            ms(times[times.len() / 2]),
            mean,
            runs.last().map_or(0, |m| m.size)
        );
        if let Some(text) = runs.last().and_then(|m| m.text.as_deref()) {
            match diff(&reference, text) {
                None => println!("    내용: GetTextFile 결과와 같음"),
                Some(d) => println!("    내용: GetTextFile 결과와 다름 — {d}"),
            }
            // 문서를 고친 뒤에도 새 내용을 읽는지(덮어쓰기·핸들 잡힌 채 읽기가 낡은 내용을 주지 않는지)
            let marker = format!("__bench_fresh_{}__", case.name().len());
            hwp.ext_text().insert(&marker)?;
            let fresh = case.run(hwp, &tmp, args.iterations + 1, &mut saved)?;
            let _ = hwp.h_action().and_then(|a| a.undo());
            let ok = fresh.text.as_deref().is_some_and(|t| t.contains(&marker));
            println!("    고친 뒤 새 내용 반영: {ok}");
        }
        // 단계별 중앙값
        let step_count = runs.first().map_or(0, |m| m.steps.len());
        for k in 0..step_count {
            let mut ts: Vec<Duration> = runs.iter().map(|m| m.steps[k].1).collect();
            ts.sort();
            println!(
                "    {:<20} {:>10} {:>10.2}",
                runs[0].steps[k].0,
                "",
                ms(ts[ts.len() / 2])
            );
        }
    }

    // 한글은 마지막으로 SaveAs한 파일의 핸들을 놓지 않으므로 GetTextFile로 핸들을 놓게 한 뒤 지운다.
    let _ = hwp.get_text_file(GetTextFileFormat::Text, false);
    for p in saved {
        let _ = std::fs::remove_file(p);
    }
    Ok(())
}
