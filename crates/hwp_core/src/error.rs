use thiserror::Error;

/// 한글(HWP) OLE 제어 중 발생할 수 있는 모든 에러를 정의합니다.
#[derive(Error, Debug)]
pub enum HwpError {
    // ---------------------------------------------------------
    // 1. 윈도우 OS / COM 시스템 에러 (가장 흔함)
    // ---------------------------------------------------------
    #[error("윈도우 COM 시스템 에러가 발생했습니다: {0}")]
    ComError(#[from] windows::core::Error),

    #[error("OLE 환경을 초기화하는 데 실패했습니다. (AfxOleInit 실패)")]
    OleInitFailed,

    #[error("한글(HWP) 객체를 생성할 수 없습니다. 한글이 설치되어 있는지 확인하세요.")]
    ConnectionFailed,

    // ---------------------------------------------------------
    // 2. 액션(Action) 및 파라미터(Parameter) 논리 에러
    // ---------------------------------------------------------
    #[error("액션 아이디 '{0}'를 찾을 수 없거나 실행할 수 없는 상태입니다.")]
    ActionNotFound(String),

    #[error("액션 '{action}' 실행에 필요한 파라미터 '{param}'가 누락되었습니다.")]
    MissingParameter { action: String, param: String },

    #[error("파라미터 '{param}'에 잘못된 타입의 값이 입력되었습니다: {details}")]
    InvalidParameterType { param: String, details: String },

    #[error("한글 명령어 실행 중 오류가 발생했습니다: {0}")]
    ExecutionFailed(String),

    // ---------------------------------------------------------
    // 3. 타입 변환 에러 (Rust <-> COM VARIANT)
    // ---------------------------------------------------------
    #[error("COM VARIANT 타입을 Rust 타입으로 변환하는 데 실패했습니다: {0}")]
    VariantConversion(String),

    #[error("추출한 텍스트가 유효한 문자열이 아닙니다.")]
    InvalidStringData,

    /// hwp ver error
    #[error("이 기능은 한글 {required_version} 이상 버전에서만 지원됩니다. (현재 버전: {current_version})")]
    UnsupportedVersion {
        required_version: String,
        current_version: String,
    },
}

/// 라이브러리 전용 Result 타입 별칭 (코드 타이핑을 획기적으로 줄여줍니다)
pub type Result<T> = std::result::Result<T, HwpError>;
