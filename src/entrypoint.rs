//! Excel12v trampoline 
//!
//! Excel exports `MdCallBack12` from the host process.  When the XLL is loaded
//! we look it up via `GetModuleHandleA` / `GetProcAddress` and cache the
//! pointer.  All calls to `excel12v` go through this cached pointer.
//!
//! This replaces the need for any outside compilation nonsense...

#![allow(non_snake_case)]

use super::types::*;
use std::sync::atomic::{AtomicPtr, Ordering};
use std::sync::Once;

// ── Win32 FFI ────────────────────────────────────────────────────────────────

type HMODULE = *mut u8;
type FARPROC = *mut u8;

extern "system" {
    fn GetModuleHandleA(lpModuleName: *const u8) -> HMODULE;
    fn GetProcAddress(hModule: HMODULE, lpProcName: *const u8) -> FARPROC;
}

// ── MdCallBack12 signature ───────────────────────────────────────────────────

/// Signature of Excel's internal callback:
///   `int PASCAL MdCallBack12(int xlfn, int coper, LPXLOPER12 *rgpxloper12, LPXLOPER12 xloper12Res)`
type Excel12Proc = unsafe extern "system" fn(
    xlfn: i32,
    coper: i32,
    rgpxloper12: *const *mut XLOPER12,
    xloper12Res: *mut XLOPER12,
) -> i32;

// ── Cached entry point ───────────────────────────────────────────────────────

static ENTRY_PT: AtomicPtr<u8> = AtomicPtr::new(std::ptr::null_mut());

static TEST_INIT: Once = Once::new();

unsafe fn free_excel_owned_oper(oper: &XLOPER12) {
    match oper.base_type() {
        XLTYPE_STR => {
            let ptr = oper.val.str_;
            if !ptr.is_null() {
                let len = *ptr as usize + 1;
                let _ = Vec::from_raw_parts(ptr, len, len);
            }
        }
        XLTYPE_MULTI => {
            let arr = &*std::ptr::addr_of!(oper.val.array);
            let total = (arr.rows * arr.columns) as usize;
            for i in 0..total {
                let elem = &*arr.lparray.add(i);
                free_excel_owned_oper(elem);
            }
            let _ = Vec::from_raw_parts(arr.lparray, total, total);
        }
        _ => {}
    }
}

unsafe extern "system" fn test_excel12(
    _xlfn: i32,
    _coper: i32,
    _rgpxloper12: *const *mut XLOPER12,
    xloper12Res: *mut XLOPER12,
) -> i32 {
    match _xlfn {
        XL_FREE => {
            if !_rgpxloper12.is_null() {
                let p = unsafe { *_rgpxloper12 };
                if !p.is_null() {
                    unsafe { free_excel_owned_oper(&*p) };
                }
            }
            XLRET_SUCCESS
        }
        XL_SHEET_NM => {
            if !xloper12Res.is_null() {
                let mut res = XLOPER12::from_str("TestSheet");
                res.xltype = XLTYPE_STR | XLBIT_XL_FREE;
                *xloper12Res = res;
            }
            XLRET_SUCCESS
        }
        XL_COERCE => {
            if _rgpxloper12.is_null() {
                return XLRET_FAILED;
            }
            let p = unsafe { *_rgpxloper12 };
            if p.is_null() {
                return XLRET_FAILED;
            }
            let base = unsafe { (*p).base_type() };
            if base != XLTYPE_SREF && base != XLTYPE_REF {
                return XLRET_FAILED;
            }
            if !xloper12Res.is_null() {
                let mut cells = vec![
                    XLOPER12::from_f64(1.0),
                    XLOPER12::from_f64(2.0),
                    XLOPER12::from_f64(3.0),
                    XLOPER12::from_f64(4.0),
                ];
                let lparray = cells.as_mut_ptr();
                std::mem::forget(cells);
                *xloper12Res = XLOPER12 {
                    val: XLOPER12Val {
                        array: std::mem::ManuallyDrop::new(XLOPER12Array {
                            lparray,
                            rows: 2,
                            columns: 2,
                        }),
                    },
                    xltype: XLTYPE_MULTI | XLBIT_XL_FREE,
                };
            }
            XLRET_SUCCESS
        }
        _ => {
            if !xloper12Res.is_null() {
                *xloper12Res = XLOPER12::from_int(0);
                unsafe {
                    (*xloper12Res).xltype |= XLBIT_XL_FREE;
                }
            }
            XLRET_SUCCESS
        }
    }
}

#[doc(hidden)]
pub fn init_test_entrypoint() {
    TEST_INIT.call_once(|| {
        ENTRY_PT.store(test_excel12 as *mut u8, Ordering::Release);
    });
}

fn fetch_entry_pt() -> Option<Excel12Proc> {
    let mut ptr = ENTRY_PT.load(Ordering::Acquire);
    if ptr.is_null() {
        unsafe {
            let hmod = GetModuleHandleA(std::ptr::null());
            if hmod.is_null() {
                return None;
            }
            ptr = GetProcAddress(hmod, b"MdCallBack12\0".as_ptr());
            if ptr.is_null() {
                return None;
            }
            ENTRY_PT.store(ptr, Ordering::Release);
        }
    }
    Some(unsafe { std::mem::transmute(ptr) })
}

// ── Public API ───────────────────────────────────────────────────────────────

/// Return codes from Excel12v.
pub const XLRET_SUCCESS: i32 = 0;
pub const XLRET_FAILED: i32 = 32;

/// Call into Excel via the `MdCallBack12` entry point.
///
///
/// # Safety
///
/// `opers` must contain `count` valid pointers.  `oper_res` must point to a
/// valid (possibly zeroed) XLOPER12.
pub unsafe fn excel12v(
    xlfn: i32,
    oper_res: *mut XLOPER12,
    count: i32,
    opers: *const *mut XLOPER12,
) -> i32 {
    match fetch_entry_pt() {
        Some(f) => f(xlfn, count, opers, oper_res),
        None => XLRET_FAILED,
    }
}

/// Convenience: call Excel12v with a slice of XLOPER12 pointers.
pub fn excel12(xlfn: i32, args: &mut [*mut XLOPER12]) -> (i32, XLOPER12) {
    let mut result = XLOPER12::nil();
    let ret = unsafe { excel12v(xlfn, &mut result, args.len() as i32, args.as_ptr()) };
    (ret, result)
}

/// Free an XLOPER12 that was returned by Excel (has `xlbitXLFree` set).
pub fn excel_free(oper: &mut XLOPER12) {
    unsafe {
        let mut p = oper as *mut XLOPER12;
        excel12v(super::types::XL_FREE, std::ptr::null_mut(), 1, &mut p);
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::types::*;

    #[test]
    fn excel12_uses_test_stub() {
        init_test_entrypoint();
        let (ret, res) = excel12(123, &mut []);
        assert_eq!(ret, XLRET_SUCCESS);
        assert_eq!(res.base_type(), XLTYPE_INT);
    }

    #[test]
    fn excel_free_is_callable_with_stub() {
        init_test_entrypoint();
        let mut oper = XLOPER12::from_int(1);
        excel_free(&mut oper);
    }
}
