//! Memory ownership helpers and xlAutoFree12 implementation.

use crate::types::*;

/// Free a DLL-owned XLOPER12 and any nested allocations.
///
/// This handles:
/// - `xltypeStr`: frees the UTF-16 buffer allocated by `XLOPER12::from_str`
/// - `xltypeMulti`: frees nested string buffers and the array itself
/// - other types: no-op
///
/// # Safety
///
/// `p` must be a valid pointer to an XLOPER12 allocated by the DLL
/// (typically via `Box::into_raw`). The memory must not be used after this call.
pub unsafe fn free_xloper_recursive(p: *mut XLOPER12) {
    if p.is_null() {
        return;
    }
    let base = (*p).xltype & 0x0FFF;
    match base {
        XLTYPE_STR => {
            let ptr = (*p).val.str_;
            if !ptr.is_null() {
                let len = *ptr as usize + 1;
                let _ = Vec::from_raw_parts(ptr, len, len);
            }
        }
        XLTYPE_MULTI => {
            let arr = &*std::ptr::addr_of!((*p).val.array);
            let total = (arr.rows * arr.columns) as usize;
            // Free strings or nested arrays inside array elements
            for i in 0..total {
                let elem_ptr = arr.lparray.add(i);
                let elem = &*elem_ptr;
                match elem.xltype & 0x0FFF {
                    XLTYPE_STR => {
                        let ptr = elem.val.str_;
                        if !ptr.is_null() {
                            let len = *ptr as usize + 1;
                            let _ = Vec::from_raw_parts(ptr, len, len);
                        }
                    }
                    XLTYPE_MULTI => {
                        free_xloper_recursive(elem_ptr);
                    }
                    _ => {}
                }
            }
            // Free the array of XLOPER12s itself
            let _ = Vec::from_raw_parts(arr.lparray, total, total);
        }
        _ => {}
    }
    // Free the XLOPER12 struct (was Box::into_raw)
    let _ = Box::from_raw(p);
}

/// Called by Excel after it copies a returned XLOPER12 whose `xltype` has
/// `xlbitDLLFree` set.  We must free all DLL-allocated memory here.
#[no_mangle]
pub extern "system" fn xlAutoFree12(p: *mut XLOPER12) {
    unsafe { free_xloper_recursive(p) };
}
