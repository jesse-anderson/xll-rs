//! XLOPER12 type definitions
//!
//! These are `#[repr(C)]` structs that match the exact memory layout Excel
//! expects. All XLOPER12 variants are included (including flow and bigdata).

#![allow(non_snake_case, non_camel_case_types, dead_code)]

use std::mem::ManuallyDrop;

// ── xltype constants ───────────────────────────────────────────────────────────

pub const XLTYPE_NUM: u32 = 0x0001;
pub const XLTYPE_STR: u32 = 0x0002;
pub const XLTYPE_BOOL: u32 = 0x0004;
pub const XLTYPE_REF: u32 = 0x0008;
pub const XLTYPE_ERR: u32 = 0x0010;
pub const XLTYPE_FLOW: u32 = 0x0020;
pub const XLTYPE_MULTI: u32 = 0x0040;
pub const XLTYPE_MISSING: u32 = 0x0080;
pub const XLTYPE_NIL: u32 = 0x0100;
pub const XLTYPE_SREF: u32 = 0x0400;
pub const XLTYPE_INT: u32 = 0x0800;
pub const XLTYPE_BIGDATA: u32 = XLTYPE_STR | XLTYPE_INT; // 0x0802

pub const XLBIT_XL_FREE: u32 = 0x1000;
pub const XLBIT_DLL_FREE: u32 = 0x4000;

// ── Excel error codes ──────────────────────────────────────────────────────────

pub const XLERR_NULL: i32 = 0;
pub const XLERR_DIV0: i32 = 7;
pub const XLERR_VALUE: i32 = 15;
pub const XLERR_REF: i32 = 23;
pub const XLERR_NAME: i32 = 29;
pub const XLERR_NUM: i32 = 36;
pub const XLERR_NA: i32 = 42;
pub const XLERR_GETTING_DATA: i32 = 43;

/// Strongly-typed Excel error code wrapper.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct XLLError(pub i32);

impl XLLError {
    pub const NULL: XLLError = XLLError(XLERR_NULL);
    pub const DIV0: XLLError = XLLError(XLERR_DIV0);
    pub const VALUE: XLLError = XLLError(XLERR_VALUE);
    pub const REF: XLLError = XLLError(XLERR_REF);
    pub const NAME: XLLError = XLLError(XLERR_NAME);
    pub const NUM: XLLError = XLLError(XLERR_NUM);
    pub const NA: XLLError = XLLError(XLERR_NA);
    pub const GETTING_DATA: XLLError = XLLError(XLERR_GETTING_DATA);

    pub fn code(self) -> i32 {
        self.0
    }
}

impl From<i32> for XLLError {
    fn from(value: i32) -> Self {
        XLLError(value)
    }
}

impl From<XLLError> for i32 {
    fn from(value: XLLError) -> Self {
        value.0
    }
}

/// Excel error with an optional custom message.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct XllError {
    code: XLLError,
    msg: Option<String>,
}

impl XllError {
    pub fn new(code: XLLError) -> Self {
        Self { code, msg: None }
    }

    pub fn msg(code: XLLError, msg: impl Into<String>) -> Self {
        Self {
            code,
            msg: Some(msg.into()),
        }
    }

    pub fn code(&self) -> i32 {
        self.code.code()
    }

    pub fn message(&self) -> Option<&str> {
        self.msg.as_deref()
    }

    fn parse_code(s: &str) -> Option<XLLError> {
        let mut key = s.trim().to_ascii_uppercase();
        if key.starts_with('#') {
            key.remove(0);
        }
        key = key.replace('?', "");
        key = key.replace('!', "");
        key = key.replace('/', "");
        key = key.replace(' ', "_");
        match key.as_str() {
            "NULL" => Some(XLLError::NULL),
            "DIV0" => Some(XLLError::DIV0),
            "VALUE" => Some(XLLError::VALUE),
            "REF" => Some(XLLError::REF),
            "NAME" => Some(XLLError::NAME),
            "NUM" => Some(XLLError::NUM),
            "NA" => Some(XLLError::NA),
            "GETTING_DATA" | "GETTINGDATA" => Some(XLLError::GETTING_DATA),
            _ => None,
        }
    }

    fn from_spec(spec: &str) -> Self {
        let trimmed = spec.trim();
        if let Some((head, tail)) = trimmed.split_once(':') {
            if let Some(code) = Self::parse_code(head) {
                let msg = tail.trim();
                return if msg.is_empty() {
                    Self::new(code)
                } else {
                    Self::msg(code, msg)
                };
            }
        }
        if let Some(code) = Self::parse_code(trimmed) {
            Self::new(code)
        } else {
            Self::new(XLLError::VALUE)
        }
    }
}

impl From<XLLError> for XllError {
    fn from(value: XLLError) -> Self {
        Self::new(value)
    }
}

impl From<i32> for XllError {
    fn from(value: i32) -> Self {
        Self::new(XLLError::from(value))
    }
}

impl From<&str> for XllError {
    fn from(value: &str) -> Self {
        Self::from_spec(value)
    }
}

impl From<String> for XllError {
    fn from(value: String) -> Self {
        Self::from_spec(&value)
    }
}

// ── Flow control codes ────────────────────────────────────────────────────────

pub const XLFLOW_HALT: u8 = 1;
pub const XLFLOW_GOTO: u8 = 2;
pub const XLFLOW_RESTART: u8 = 8;
pub const XLFLOW_PAUSE: u8 = 16;
pub const XLFLOW_RESUME: u8 = 64;

// ── Excel C API function numbers ───────────────────────────────────────────────

const XL_SPECIAL: i32 = 0x4000;

pub const XL_FREE: i32 = 0 | XL_SPECIAL;
pub const XL_GET_NAME: i32 = 9 | XL_SPECIAL;
pub const XLF_REGISTER: i32 = 149;
pub const XL_COERCE: i32 = 2 | XL_SPECIAL;
pub const XL_SHEET_ID: i32 = 4 | XL_SPECIAL;
pub const XL_SHEET_NM: i32 = 5 | XL_SPECIAL;
pub const XL_GET_INST: i32 = 7 | XL_SPECIAL;
pub const XL_GET_INST_PTR: i32 = 19 | XL_SPECIAL; // Excel 2010+ (BigData handle)

// ── XLOPER12 ───────────────────────────────────────────────────────────────────

/// The core XLOPER12 struct.
#[repr(C)]
pub struct XLOPER12 {
    pub val: XLOPER12Val,
    pub xltype: u32,
}

/// Union of all XLOPER12 value types.
///
/// We use `ManuallyDrop` wrappers for struct variants to allow them in a
/// `Copy` union without implementing `Drop` on the union itself (memory
/// is managed externally via `xlAutoFree12`).
#[repr(C)]
pub union XLOPER12Val {
    pub num: f64,
    pub str_: *mut u16,
    pub xbool: i32,
    pub err: i32,
    pub w: i32,
    pub array: ManuallyDrop<XLOPER12Array>,
    pub sref: ManuallyDrop<XLOPER12SRef>,
    pub mref: ManuallyDrop<XLOPER12MRef>,
    pub flow: ManuallyDrop<XLOPER12Flow>,
    pub bigdata: ManuallyDrop<XLOPER12BigData>,
}

#[repr(C)]
#[derive(Clone, Copy)]
pub struct XLOPER12Array {
    pub lparray: *mut XLOPER12,
    pub rows: i32,
    pub columns: i32,
}

#[repr(C)]
#[derive(Clone, Copy)]
pub struct XLOPER12SRef {
    pub count: u16,
    pub ref_: XLREF12,
}

#[repr(C)]
#[derive(Clone, Copy)]
pub struct XLREF12 {
    pub rw_first: i32,
    pub rw_last: i32,
    pub col_first: i32,
    pub col_last: i32,
}

#[repr(C)]
#[derive(Clone, Copy)]
pub struct XLOPER12MRef {
    pub lpmref: *mut XLMREF12,
    pub id_sheet: usize,
}

#[repr(C)]
pub struct XLMREF12 {
    pub count: u16,
    pub reftbl: [XLREF12; 1],
}

#[repr(C)]
#[derive(Clone, Copy)]
pub union XLOPER12FlowVal {
    pub level: i32,      // xlflowRestart
    pub tbctrl: i32,     // xlflowPause
    pub id_sheet: usize, // xlflowGoto (IDSHEET)
}

#[repr(C)]
#[derive(Clone, Copy)]
pub struct XLOPER12Flow {
    pub valflow: XLOPER12FlowVal,
    pub rw: i32,
    pub col: i32,
    pub xlflow: u8,
}

#[repr(C)]
#[derive(Clone, Copy)]
pub union XLOPER12BigDataHandle {
    pub lpb_data: *mut u8,  // data passed to Excel
    pub hdata: *mut u8,     // HANDLE returned from Excel
}

#[repr(C)]
#[derive(Clone, Copy)]
pub struct XLOPER12BigData {
    pub h: XLOPER12BigDataHandle,
    pub cb_data: i32,
}

// ── Convenience constructors ───────────────────────────────────────────────────

impl XLOPER12 {
    /// Create a nil (empty) XLOPER12.
    pub fn nil() -> Self {
        Self {
            val: XLOPER12Val { w: 0 },
            xltype: XLTYPE_NIL,
        }
    }

    /// Create a numeric XLOPER12.
    pub fn from_f64(v: f64) -> Self {
        Self {
            val: XLOPER12Val { num: v },
            xltype: XLTYPE_NUM,
        }
    }

    /// Create an integer XLOPER12.
    pub fn from_int(v: i32) -> Self {
        Self {
            val: XLOPER12Val { w: v },
            xltype: XLTYPE_INT,
        }
    }

    /// Create a boolean XLOPER12.
    pub fn from_bool(v: bool) -> Self {
        Self {
            val: XLOPER12Val { xbool: if v { 1 } else { 0 } },
            xltype: XLTYPE_BOOL,
        }
    }

    /// Create an error XLOPER12.
    pub fn from_err(code: i32) -> Self {
        Self {
            val: XLOPER12Val { err: code },
            xltype: XLTYPE_ERR,
        }
    }

    /// Create a missing-argument XLOPER12.
    pub fn missing() -> Self {
        Self {
            val: XLOPER12Val { w: 0 },
            xltype: XLTYPE_MISSING,
        }
    }

    /// Create a string XLOPER12 from a Rust `&str`.
    ///
    /// Allocates a length-counted UTF-16 buffer: `[len, char0, char1, ...]`.
    /// Sets `xlbitDLLFree` so the memory is reclaimed in `xlAutoFree12`.
    pub fn from_str(s: &str) -> Self {
        let utf16: Vec<u16> = s.encode_utf16().collect();
        let len = utf16.len();
        if len > 32767 {
            return Self::from_err(XLERR_VALUE);
        }

        // Build pascal-style string: length prefix followed by chars
        let mut buf: Vec<u16> = Vec::with_capacity(len + 1);
        buf.push(len as u16);
        buf.extend_from_slice(&utf16);

        let ptr = buf.as_mut_ptr();
        std::mem::forget(buf);

        Self {
            val: XLOPER12Val { str_: ptr },
            xltype: XLTYPE_STR | XLBIT_DLL_FREE,
        }
    }

    /// Extract a Rust `String` from this XLOPER12 if it is a string type.
    pub fn as_string(&self) -> Option<String> {
        if (self.xltype & 0x0FFF) != XLTYPE_STR {
            return None;
        }
        unsafe {
            let ptr = self.val.str_;
            if ptr.is_null() {
                return None;
            }
            let len = *ptr as usize;
            let slice = std::slice::from_raw_parts(ptr.add(1), len);
            Some(String::from_utf16_lossy(slice))
        }
    }

    /// Extract an `f64` from this XLOPER12 if it is numeric or integer type.
    pub fn as_f64(&self) -> Option<f64> {
        match self.xltype & 0x0FFF {
            XLTYPE_NUM => Some(unsafe { self.val.num }),
            XLTYPE_INT => Some(unsafe { self.val.w } as f64),
            _ => None,
        }
    }

    /// Extract a `bool` from this XLOPER12 if it is a boolean type.
    pub fn as_bool(&self) -> Option<bool> {
        match self.xltype & 0x0FFF {
            XLTYPE_BOOL => Some(unsafe { self.val.xbool } != 0),
            _ => None,
        }
    }

    /// Returns `true` if this XLOPER12 represents a missing argument.
    pub fn is_missing(&self) -> bool {
        (self.xltype & 0x0FFF) == XLTYPE_MISSING
    }

    /// Base type with memory flags masked off.
    pub fn base_type(&self) -> u32 {
        self.xltype & 0x0FFF
    }
}
