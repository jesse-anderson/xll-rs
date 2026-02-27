//! Helpers for returning XLOPER12 values from UDFs without leaking.
//!
//! This wrapper always sets `xlbitDLLFree` on returned values so Excel will
//! call `xlAutoFree12` after copying the result.
//!
//! Limitations:
//! - This only enforces the contract if you use `XlReturn`. It cannot stop
//!   callers from manually allocating an `XLOPER12` and forgetting the flag.
//! - Scalar returns still incur a heap allocation. For tiny UDFs this adds
//!   overhead, but it's correct and safe by default.
//!
//! Optimization path:
//! - An `xll-gen` wrapper should default to `xlbitDLLFree`, but allow opt-out
//!   for known-safe fast paths (stack/static returns or cached allocations).

use crate::types::*;

/// Owned return value that guarantees `xlbitDLLFree` is set.
pub struct XlReturn(XLOPER12);

impl XlReturn {
    /// Wrap an existing XLOPER12 as a DLL-owned return value.
    pub fn from_oper(mut oper: XLOPER12) -> Self {
        oper.xltype |= XLBIT_DLL_FREE;
        Self(oper)
    }

    pub fn num(v: f64) -> Self {
        Self::from_oper(XLOPER12::from_f64(v))
    }

    pub fn int(v: i32) -> Self {
        Self::from_oper(XLOPER12::from_int(v))
    }

    pub fn bool(v: bool) -> Self {
        Self::from_oper(XLOPER12::from_bool(v))
    }

    pub fn str(s: &str) -> Self {
        Self::from_oper(XLOPER12::from_str(s))
    }

    pub fn err(code: i32) -> Self {
        Self::from_oper(XLOPER12::from_err(code))
    }

    pub fn nil() -> Self {
        Self::from_oper(XLOPER12::nil())
    }

    pub fn missing() -> Self {
        Self::from_oper(XLOPER12::missing())
    }

    /// Convert into a raw pointer for Excel to consume.
    pub fn into_raw(self) -> *mut XLOPER12 {
        Box::into_raw(Box::new(self.0))
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::memory::xlAutoFree12;

    #[test]
    fn xlreturn_sets_dll_free_flag() {
        let p = XlReturn::num(3.5).into_raw();
        unsafe {
            assert_ne!((*p).xltype & XLBIT_DLL_FREE, 0);
            assert_eq!((*p).base_type(), XLTYPE_NUM);
        }
        xlAutoFree12(p);
    }

    #[test]
    fn xlreturn_string_roundtrip() {
        let p = XlReturn::str("hi").into_raw();
        unsafe {
            assert_eq!((*p).as_string().as_deref(), Some("hi"));
        }
        xlAutoFree12(p);
    }

    #[test]
    fn xlreturn_nil_and_missing() {
        let p = XlReturn::nil().into_raw();
        unsafe {
            assert_eq!((*p).base_type(), XLTYPE_NIL);
        }
        xlAutoFree12(p);

        let p = XlReturn::missing().into_raw();
        unsafe {
            assert_eq!((*p).base_type(), XLTYPE_MISSING);
        }
        xlAutoFree12(p);
    }
}
