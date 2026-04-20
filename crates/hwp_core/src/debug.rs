use std::fs::File;
use std::io::Write;
use std::sync::{Mutex, OnceLock};

static LOG_FILE: OnceLock<Option<Mutex<File>>> = OnceLock::new();

fn get_log_file() -> Option<&'static Mutex<File>> {
    LOG_FILE.get_or_init(|| {
        let local_app_data = std::env::var("LOCALAPPDATA").ok()?;
        let dir = std::path::Path::new(&local_app_data).join("HwpAddon");
        std::fs::create_dir_all(&dir).ok()?;
        let path = dir.join("hwp_addon_debug.log");
        let file = std::fs::OpenOptions::new()
            .create(true)
            .write(true)
            .truncate(true)
            .open(path)
            .ok()?;
        Some(Mutex::new(file))
    }).as_ref()
}

/// 디버그 로그를 파일에 기록합니다.
pub fn log(tag: &str, msg: &str) {
    let Some(f) = get_log_file() else { return };
    if let Ok(mut f) = f.lock() {
        let now = chrono::Local::now().format("%Y-%m-%d %H:%M:%S%.3f");
        let _ = writeln!(f, "[{now}][{tag}] {msg}");
    }
}
