use std::path::Path;
use std::sync::OnceLock;
use windows::Win32::Graphics::Gdi::{
    CreateDIBitmap, GetDC, ReleaseDC, BITMAPINFO, CBM_INIT, DIB_RGB_COLORS,
};
use windows::Win32::UI::WindowsAndMessaging::{LoadImageW, IMAGE_BITMAP, LR_LOADFROMFILE};

/// HWP 툴바용 비트맵 스트립을 로드하고 보관합니다.
///
/// HWP 요구사항: 960×16 (16px × 60), 32bit RGBA BMP
///
/// # Example
/// ```ignore
/// use hwp_addon::toolbar::ToolbarBitmap;
///
/// static TOOLBAR: ToolbarBitmap = ToolbarBitmap::new();
///
/// // get_action_image 구현 안에서:
/// TOOLBAR.load("toolbar.bmp"); // 첫 호출 시만 로드
/// TOOLBAR.image(0)             // index 0번 아이콘
/// ```
pub struct ToolbarBitmap {
    inner: OnceLock<Option<isize>>,
}

// HBITMAP 핸들은 프로세스 전역 GDI 리소스이며 DLL 수명 동안 유지되므로 안전합니다.
unsafe impl Sync for ToolbarBitmap {}
unsafe impl Send for ToolbarBitmap {}

impl ToolbarBitmap {
    pub const fn new() -> Self {
        Self {
            inner: OnceLock::new(),
        }
    }

    /// BMP 파일을 로드합니다. 이미 로드되었으면 아무 일도 하지 않습니다.
    pub fn load(&self, path: impl AsRef<Path>) {
        self.inner.get_or_init(|| {
            let path = path.as_ref();
            let wide: Vec<u16> = path
                .as_os_str()
                .encode_wide()
                .chain(std::iter::once(0))
                .collect();
            let handle = unsafe {
                LoadImageW(
                    None,
                    windows::core::PCWSTR(wide.as_ptr()),
                    IMAGE_BITMAP,
                    0,
                    0,
                    LR_LOADFROMFILE,
                )
                .ok()?
            };
            Some(handle.0 as isize)
        });
    }

    /// 컴파일 시 포함된 BMP 바이트 슬라이스에서 비트맵을 로드합니다.
    /// 이미 로드되었으면 아무 일도 하지 않습니다.
    pub fn load_from_bytes(&self, data: &[u8]) {
        self.inner.get_or_init(|| bmp_to_hbitmap(data));
    }

    /// `get_action_image`에서 사용할 `(hBitmap, imageIndex)` 튜플을 반환합니다.
    /// 비트맵이 로드되지 않았으면 `None`을 반환합니다.
    pub fn image(&self, index: i32) -> Option<(isize, i32)> {
        let handle = (*self.inner.get()?)?;
        Some((handle, index))
    }
}

impl Default for ToolbarBitmap {
    fn default() -> Self {
        Self::new()
    }
}

fn bmp_to_hbitmap(data: &[u8]) -> Option<isize> {
    if data.len() < 54 || &data[0..2] != b"BM" {
        return None;
    }
    let pixel_offset = u32::from_le_bytes(data[10..14].try_into().ok()?) as usize;
    if pixel_offset >= data.len() {
        return None;
    }
    unsafe {
        let hdc = GetDC(None);
        let bmi = data.as_ptr().add(14) as *const BITMAPINFO;
        let bits = data.as_ptr().add(pixel_offset).cast();
        let hbm = CreateDIBitmap(
            hdc,
            Some(&(*bmi).bmiHeader),
            CBM_INIT as u32,
            Some(bits),
            Some(bmi),
            DIB_RGB_COLORS,
        );
        ReleaseDC(None, hdc);
        if hbm.is_invalid() {
            None
        } else {
            Some(hbm.0 as isize)
        }
    }
}

use std::os::windows::ffi::OsStrExt;
