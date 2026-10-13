//! 애드온 DLL용 파일 로그 구독자 (`file-log` 기능).
//!
//! `hwp_core`와 `hwp_addon`은 [`tracing`] 이벤트를 내보내기만 하고 구독자를
//! 설치하지 않으므로, 아무것도 하지 않으면 로그는 어디에도 기록되지 않습니다.
//! 애드온에서 로그 파일이 필요하면 [`HwpUserAction::on_load`]에서
//! [`init_default`] 또는 [`init_file`]을 호출합니다.
//!
//! ```ignore
//! fn on_load(&self, hwp: &HwpObject) -> Result<bool> {
//!     let _ = hwp_addon::logging::init_default(hwp_addon::logging::LevelFilter::DEBUG);
//!     self.setup_toolbar(hwp)
//! }
//! ```
//!
//! 기록은 호출한 스레드에서 동기로 이뤄집니다. `tracing-appender`의 non_blocking
//! 같은 백그라운드 writer 스레드는 HWP가 DLL을 언로드할 때 언로드된 코드를
//! 실행하다 크래시를 낼 수 있으므로 쓰지 않습니다.
//!
//! [`HwpUserAction::on_load`]: crate::hwp_user_action::HwpUserAction::on_load

use std::fs::{File, OpenOptions};
use std::io;
use std::path::{Path, PathBuf};
use std::sync::Mutex;

pub use tracing_subscriber::filter::LevelFilter;
use tracing_subscriber::fmt::time::ChronoLocal;

/// 구독자 설치에 성공했는지 여부. 실패한 경우 다음 호출에서 다시 시도한다.
static INSTALLED: Mutex<bool> = Mutex::new(false);

/// 기본 로그 경로 `%LOCALAPPDATA%\HwpAddon\hwp_addon_debug.log`를 돌려줍니다.
///
/// `LOCALAPPDATA` 환경변수가 없으면 `None`입니다.
pub fn default_log_path() -> Option<PathBuf> {
    let local_app_data = std::env::var_os("LOCALAPPDATA")?;
    Some(
        Path::new(&local_app_data)
            .join("HwpAddon")
            .join("hwp_addon_debug.log"),
    )
}

/// [`default_log_path`]에 로그를 기록하는 구독자를 설치합니다.
///
/// 동작은 [`init_file`]과 같습니다.
pub fn init_default(level: LevelFilter) -> io::Result<()> {
    let path = default_log_path()
        .ok_or_else(|| io::Error::new(io::ErrorKind::NotFound, "LOCALAPPDATA 없음"))?;
    init_file(path, level)
}

/// `path`에 `level` 이상의 로그를 기록하는 전역 구독자를 설치합니다.
///
/// 상위 디렉터리가 없으면 만들고, 파일은 처음 호출될 때 비웁니다.
/// `on_load`처럼 여러 번 불리는 곳에서 호출해도 되도록, 두 번째 호출부터는
/// 아무것도 하지 않고 `Ok(())`를 돌려줍니다. 설치에 실패했으면 다음 호출에서
/// 다시 시도합니다. 이 DLL 안에 다른 전역 구독자가 이미 있으면 에러입니다.
pub fn init_file(path: impl AsRef<Path>, level: LevelFilter) -> io::Result<()> {
    let mut installed = INSTALLED.lock().unwrap_or_else(|e| e.into_inner());
    if !*installed {
        install(path.as_ref(), level)?;
        *installed = true;
    }
    Ok(())
}

fn install(path: &Path, level: LevelFilter) -> io::Result<()> {
    if let Some(dir) = path.parent() {
        std::fs::create_dir_all(dir)?;
    }
    let file: File = OpenOptions::new()
        .create(true)
        .write(true)
        .truncate(true)
        .open(path)?;

    tracing_subscriber::fmt()
        .with_writer(Mutex::new(file))
        .with_max_level(level)
        .with_timer(ChronoLocal::new("%Y-%m-%d %H:%M:%S%.3f".to_owned()))
        .with_ansi(false)
        .try_init()
        .map_err(io::Error::other)
}
