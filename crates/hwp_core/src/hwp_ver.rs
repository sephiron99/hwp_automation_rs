#[derive(Debug, Clone, PartialEq, Eq)]
pub enum HwpVer {
    V2018(u8),     // u8는 minor version
    V2022(u8),     // u8는 minor version
    V2024(u8),     // u8는 minor version
    Other(u8, u8), // (major, minor)
}

impl HwpVer {
    /// 32비트 정수형 버전 코드를 바이트 단위로 쪼개어 파싱합니다.
    pub fn from_u32(version_code: u32) -> Self {
        let bytes: [u8; 4] = version_code.to_le_bytes();

        let hwp_major: u8 = bytes[3];
        let hwp_minor: u8 = bytes[2];
        let _ocx_major: u8 = bytes[1];
        let _ocx_minor: u8 = bytes[0];

        // 파싱된 주 버전(Major) 숫자를 바탕으로 Enum을 생성합니다.
        match hwp_major {
            10 => Self::V2018(hwp_minor),
            12 => Self::V2022(hwp_minor),
            13 => Self::V2024(hwp_minor),
            _ => Self::Other(hwp_major, hwp_minor),
        }
    }

    /// 한글의 내부 버전(Major) 숫자로 매핑하여 대소 비교를 수행합니다.
    /// (HWP 2018 = 10, HWP 2022 = 12, HWP 2024 = 13)
    /// (major, minor)
    fn as_number(&self) -> (u8, u8) {
        match self {
            HwpVer::V2018(minor_v) => (10, *minor_v),
            HwpVer::V2022(minor_v) => (12, *minor_v),
            HwpVer::V2024(minor_v) => (13, *minor_v),
            HwpVer::Other(major_v, minor_v) => (*major_v, *minor_v),
        }
    }

    /// 현재 버전이 요구 버전(required) 이상인지 확인합니다.
    pub fn is_at_least(&self, required: &HwpVer) -> bool {
        self.as_number() >= required.as_number()
    }

    /// "10, 0, 0, 14727" 또는 "10.0.0.14727" 형태의 버전 문자열에서 파싱합니다.
    pub fn from_version_string(v: &str) -> Self {
        // "10, 0, 0, 14727" → 첫 번째 숫자(major)와 두 번째 숫자(minor) 추출
        let mut parts = v
            .split([',', '.'])
            .filter_map(|s| s.trim().parse::<u8>().ok());
        let major = parts.next().unwrap_or(0);
        let minor = parts.next().unwrap_or(0);
        match major {
            10 => Self::V2018(minor),
            12 => Self::V2022(minor),
            13 => Self::V2024(minor),
            _ => Self::Other(major, minor),
        }
    }

    /// 에러 메시지 출력을 위한 보기 좋은 문자열 반환
    pub fn display_name(&self) -> String {
        match self {
            HwpVer::V2018(_) => "한글 2018".to_string(),
            HwpVer::V2022(_) => "한글 2022".to_string(),
            HwpVer::V2024(_) => "한글 2024".to_string(),
            HwpVer::Other(_, _) => "알 수 없는 버전()".to_string(),
        }
    }
}
