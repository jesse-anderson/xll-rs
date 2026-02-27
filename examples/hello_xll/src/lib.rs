use xll_rs::convert::return_xl_error;
use xll_rs::register::Reg;
use xll_rs::returning::XlReturn;
use xll_rs::types::*;

#[no_mangle]
pub extern "system" fn xl_hello() -> *mut XLOPER12 {
    XlReturn::str("Hello from Rust").into_raw()
}

#[no_mangle]
pub extern "system" fn xl_add(a: *const XLOPER12, b: *const XLOPER12) -> *mut XLOPER12 {
    if a.is_null() || b.is_null() {
        return return_xl_error(XLERR_VALUE);
    }
    unsafe {
        let av = (*a).as_f64().unwrap_or(0.0);
        let bv = (*b).as_f64().unwrap_or(0.0);
        XlReturn::num(av + bv).into_raw()
    }
}

#[no_mangle]
pub extern "system" fn xlAutoOpen() -> i32 {
    let reg = Reg::new();
    reg.add(
        "xl_hello",
        "Q$",
        "HELLO.RUST",
        "",
        "xll-rs",
        "Returns a greeting string",
        &[],
    );
    reg.add(
        "xl_add",
        "QQQ$",
        "ADD.RUST",
        "a, b",
        "xll-rs",
        "Adds two numbers",
        &["First number", "Second number"],
    );
    1
}

#[no_mangle]
pub extern "system" fn xlAutoClose() -> i32 {
    1
}

#[no_mangle]
pub extern "system" fn xlAddInManagerInfo12(action: *const XLOPER12) -> *mut XLOPER12 {
    if !action.is_null() {
        let oper = unsafe { &*action };
        let is_one = match oper.base_type() {
            XLTYPE_NUM => (unsafe { oper.val.num }) == 1.0,
            XLTYPE_INT => (unsafe { oper.val.w }) == 1,
            _ => false,
        };
        if is_one {
            return XlReturn::str("xll-rs example").into_raw();
        }
    }
    return_xl_error(XLERR_VALUE)
}

// Excel calls this after it copies results with xlbitDLLFree set
pub use xll_rs::memory::xlAutoFree12;
