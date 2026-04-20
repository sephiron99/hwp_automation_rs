/// 한글 OLE 클라이언트 예제
///
/// Running Object Table에서 실행 중인 한글 인스턴스를 열거하고
/// 각 인스턴스의 버전·경로·창 정보를 출력합니다.
/// 또한 각 창에 창 번호를 순서대로 입력합니다.
fn main() -> hwp_com::Result<()> {
    // ROT 열거 대상이 되도록 먼저 한글 인스턴스를 하나 띄워 활성화 창을 표시합니다.
    let hwp = hwp_com::HwpClient::new()?;
    hwp.windows()?.active()?.set_visible(true)?;

    dump_running_instances()?;
    label_all_windows()?;
    Ok(())
}

/// 실행 중인 모든 한글 인스턴스의 각 창(도큐먼트)을 순서대로 활성화하며,
/// 해당 창에 자신의 인덱스를 텍스트로 삽입합니다.
///
/// `XHwpDocuments.Item(i)` (0-based)로 각 문서를 얻어 `SetActive_XHwpDocument`로
/// 활성화한 뒤, `HwpObject.insert_text("Window <i>")`를 호출합니다.
fn label_all_windows() -> hwp_com::Result<()> {
    let instances = hwp_com::HwpClient::list_running()?;
    for (_moniker, hwp) in &instances {
        let docs = match hwp.documents() {
            Ok(d) => d,
            Err(_) => continue,
        };
        let count = docs.count().unwrap_or(0);
        for i in 0..count {
            let Ok(doc) = docs.item(i) else { continue };
            if doc.set_active().is_err() {
                continue;
            }
            let _ = hwp.insert_text(&format!("Hello Hwp, Window {i}"));
        }
    }
    Ok(())
}

/// Running Object Table에 등록된 모든 한글 OLE 인스턴스를 조회하고,
/// 각 인스턴스의 버전·경로·창 목록(가시성·좌표·크기)을 표준출력으로 출력합니다.
fn dump_running_instances() -> hwp_com::Result<()> {
    let instances = hwp_com::HwpClient::list_running()?;
    println!("== 실행 중인 한글 OLE 인스턴스: {}개 ==", instances.len());

    for (idx, (moniker, hwp)) in instances.iter().enumerate() {
        println!();
        println!("[{idx}] moniker = {moniker}");
        println!("    version = {}", hwp.version().display_name());
        match hwp.path() {
            Ok(p) if !p.is_empty() => println!("    path    = {p}"),
            Ok(_) => println!("    path    = (저장되지 않음)"),
            Err(e) => println!("    path    = <오류: {e}>"),
        }

        let windows = match hwp.windows() {
            Ok(w) => w,
            Err(e) => {
                println!("    windows = <오류: {e}>");
                continue;
            }
        };
        let count = windows.count().unwrap_or(0);
        println!("    창 {count}개");

        // Item은 0-based (count=2일 때 유효 인덱스는 0, 1)
        for i in 0..count {
            match windows.item(i) {
                Ok(w) => {
                    let visible = w.visible().unwrap_or(false);
                    let left = w.left().unwrap_or(-1);
                    let top = w.top().unwrap_or(-1);
                    let width = w.width().unwrap_or(-1);
                    let height = w.height().unwrap_or(-1);
                    println!(
                        "      [{i}] visible={visible} pos=({left},{top}) size=({width}x{height})"
                    );
                }
                Err(e) => println!("      [{i}] <오류: {e}>"),
            }
        }
    }

    Ok(())
}
