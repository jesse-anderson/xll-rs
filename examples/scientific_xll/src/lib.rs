use xll_rs::convert::{build_multi, return_xl_error, xloper_to_columns, xloper_to_f64_vec};
use xll_rs::entrypoint::{excel12, excel_free, XLRET_SUCCESS};
use xll_rs::register::Reg;
use xll_rs::returning::XlReturn;
use xll_rs::types::*;

use std::mem::ManuallyDrop;
use std::sync::OnceLock;

static REG_LOG: OnceLock<String> = OnceLock::new();

// ── helpers ────────────────────────────────────────────────────────────────

fn parse_optional_f64(p: *const XLOPER12, default: f64) -> Result<f64, i32> {
    if p.is_null() {
        return Err(XLERR_VALUE);
    }
    let oper = unsafe { &*p };
    match oper.base_type() {
        XLTYPE_MISSING | XLTYPE_NIL => Ok(default),
        XLTYPE_NUM => Ok(unsafe { oper.val.num }),
        XLTYPE_INT => Ok(unsafe { oper.val.w } as f64),
        XLTYPE_ERR => Err(unsafe { oper.val.err }),
        _ => Err(XLERR_VALUE),
    }
}

fn parse_optional_bool(p: *const XLOPER12, default: bool) -> Result<bool, i32> {
    if p.is_null() {
        return Err(XLERR_VALUE);
    }
    let oper = unsafe { &*p };
    match oper.base_type() {
        XLTYPE_MISSING | XLTYPE_NIL => Ok(default),
        XLTYPE_BOOL => Ok(unsafe { oper.val.xbool } != 0),
        XLTYPE_NUM => Ok(unsafe { oper.val.num } != 0.0),
        XLTYPE_INT => Ok(unsafe { oper.val.w } != 0),
        XLTYPE_ERR => Err(unsafe { oper.val.err }),
        _ => Err(XLERR_VALUE),
    }
}

fn build_kv_table(rows: &[(&str, f64)]) -> *mut XLOPER12 {
    let n_rows = rows.len();
    let mut cells: Vec<XLOPER12> = Vec::with_capacity(n_rows * 2);
    for (label, value) in rows {
        cells.push(XLOPER12::from_str(label));
        cells.push(XLOPER12::from_f64(*value));
    }
    build_multi(cells, n_rows, 2)
}

fn sheet_name_from_ref(p: *const XLOPER12) -> Option<String> {
    if p.is_null() {
        return None;
    }
    let mut args = [p as *mut XLOPER12];
    let (ret, mut res) = excel12(XL_SHEET_NM, &mut args);
    if ret != XLRET_SUCCESS {
        return None;
    }
    let name = if res.base_type() == XLTYPE_STR {
        res.as_string()
    } else {
        None
    };
    if (res.xltype & XLBIT_XL_FREE) != 0 {
        excel_free(&mut res);
    }
    name
}

fn clone_excel_oper(oper: &XLOPER12) -> Result<*mut XLOPER12, i32> {
    match oper.base_type() {
        XLTYPE_NUM => Ok(XlReturn::num(unsafe { oper.val.num }).into_raw()),
        XLTYPE_INT => Ok(XlReturn::int(unsafe { oper.val.w }).into_raw()),
        XLTYPE_BOOL => Ok(XlReturn::bool(unsafe { oper.val.xbool } != 0).into_raw()),
        XLTYPE_ERR => Ok(XlReturn::err(unsafe { oper.val.err }).into_raw()),
        XLTYPE_NIL => Ok(XlReturn::nil().into_raw()),
        XLTYPE_STR => {
            let s = oper.as_string().ok_or(XLERR_VALUE)?;
            Ok(XlReturn::str(&s).into_raw())
        }
        XLTYPE_MULTI => {
            let arr = unsafe { &*std::ptr::addr_of!(oper.val.array) };
            let rows = arr.rows as usize;
            let cols = arr.columns as usize;
            let total = rows * cols;
            let mut cells: Vec<XLOPER12> = Vec::with_capacity(total);
            for i in 0..total {
                let elem = unsafe { &*arr.lparray.add(i) };
                match elem.base_type() {
                    XLTYPE_NUM => cells.push(XLOPER12::from_f64(unsafe { elem.val.num })),
                    XLTYPE_INT => cells.push(XLOPER12::from_int(unsafe { elem.val.w })),
                    XLTYPE_BOOL => cells.push(XLOPER12::from_bool(unsafe { elem.val.xbool } != 0)),
                    XLTYPE_ERR => cells.push(XLOPER12::from_err(unsafe { elem.val.err })),
                    XLTYPE_NIL => cells.push(XLOPER12::nil()),
                    XLTYPE_STR => {
                        let s = elem.as_string().ok_or(XLERR_VALUE)?;
                        cells.push(XLOPER12::from_str(&s));
                    }
                    _ => return Err(XLERR_VALUE),
                }
            }
            Ok(build_multi(cells, rows, cols))
        }
        _ => Err(XLERR_VALUE),
    }
}

fn coerce_to_owned(oper: &XLOPER12) -> Result<*mut XLOPER12, i32> {
    let mut target_type = XLOPER12::from_int(XLTYPE_MULTI as i32);
    let mut args = [
        oper as *const XLOPER12 as *mut XLOPER12,
        &mut target_type as *mut XLOPER12,
    ];
    let (ret, mut res) = excel12(XL_COERCE, &mut args);
    if ret != XLRET_SUCCESS {
        return Err(XLERR_VALUE);
    }
    let out = clone_excel_oper(&res)?;
    if (res.xltype & XLBIT_XL_FREE) != 0 {
        excel_free(&mut res);
    }
    Ok(out)
}

// ── UDFs ────────────────────────────────────────────────────────────────────

#[no_mangle]
pub extern "system" fn xl_sci_version() -> *mut XLOPER12 {
    XlReturn::str(env!("CARGO_PKG_VERSION")).into_raw()
}

#[no_mangle]
pub extern "system" fn xl_sci_hello() -> *mut XLOPER12 {
    XlReturn::str("Hello from xll-rs").into_raw()
}

#[no_mangle]
pub extern "system" fn xl_sci_add(a: *const XLOPER12, b: *const XLOPER12) -> *mut XLOPER12 {
    if a.is_null() || b.is_null() {
        return return_xl_error(XLERR_VALUE);
    }
    let av = unsafe { &*a }.as_f64().unwrap_or(0.0);
    let bv = unsafe { &*b }.as_f64().unwrap_or(0.0);
    XlReturn::num(av + bv).into_raw()
}

#[no_mangle]
pub extern "system" fn xl_sci_not(p: *const XLOPER12) -> *mut XLOPER12 {
    if p.is_null() {
        return return_xl_error(XLERR_VALUE);
    }
    let v = unsafe { &*p }.as_bool().unwrap_or(false);
    XlReturn::bool(!v).into_raw()
}

#[no_mangle]
pub extern "system" fn xl_sci_echo(p: *const XLOPER12) -> *mut XLOPER12 {
    if p.is_null() {
        return return_xl_error(XLERR_VALUE);
    }
    let oper = unsafe { &*p };
    if let Some(s) = oper.as_string() {
        XlReturn::str(&s).into_raw()
    } else {
        return_xl_error(XLERR_VALUE)
    }
}

#[no_mangle]
pub extern "system" fn xl_sci_mean(y_range: *const XLOPER12) -> *mut XLOPER12 {
    let y = match xloper_to_f64_vec(y_range) {
        Ok(v) => v,
        Err(code) => return return_xl_error(code),
    };
    if y.is_empty() {
        return return_xl_error(XLERR_VALUE);
    }
    let sum: f64 = y.iter().sum();
    XlReturn::num(sum / y.len() as f64).into_raw()
}

#[no_mangle]
pub extern "system" fn xl_sci_describe(y_range: *const XLOPER12) -> *mut XLOPER12 {
    let y = match xloper_to_f64_vec(y_range) {
        Ok(v) => v,
        Err(code) => return return_xl_error(code),
    };
    if y.is_empty() {
        return return_xl_error(XLERR_VALUE);
    }
    let n = y.len() as f64;
    let mean = y.iter().sum::<f64>() / n;
    let min = y.iter().cloned().fold(f64::INFINITY, f64::min);
    let max = y.iter().cloned().fold(f64::NEG_INFINITY, f64::max);
    let var = y.iter().map(|v| (v - mean) * (v - mean)).sum::<f64>() / (n - 1.0).max(1.0);
    let std = var.sqrt();

    build_kv_table(&[
        ("Mean", mean),
        ("Std Dev", std),
        ("Min", min),
        ("Max", max),
    ])
}

#[no_mangle]
pub extern "system" fn xl_sci_corr(
    y_range: *const XLOPER12,
    x_range: *const XLOPER12,
) -> *mut XLOPER12 {
    let y = match xloper_to_f64_vec(y_range) {
        Ok(v) => v,
        Err(code) => return return_xl_error(code),
    };
    let x = match xloper_to_f64_vec(x_range) {
        Ok(v) => v,
        Err(code) => return return_xl_error(code),
    };
    if y.len() != x.len() || y.is_empty() {
        return return_xl_error(XLERR_VALUE);
    }
    let n = y.len() as f64;
    let mean_y = y.iter().sum::<f64>() / n;
    let mean_x = x.iter().sum::<f64>() / n;
    let mut num = 0.0;
    let mut den_y = 0.0;
    let mut den_x = 0.0;
    for i in 0..y.len() {
        let dy = y[i] - mean_y;
        let dx = x[i] - mean_x;
        num += dy * dx;
        den_y += dy * dy;
        den_x += dx * dx;
    }
    let denom = (den_y * den_x).sqrt();
    if denom == 0.0 {
        return return_xl_error(XLERR_NUM);
    }
    XlReturn::num(num / denom).into_raw()
}

#[no_mangle]
pub extern "system" fn xl_sci_colsum(x_range: *const XLOPER12) -> *mut XLOPER12 {
    let (cols, _n_rows, n_cols) = match xloper_to_columns(x_range) {
        Ok(v) => v,
        Err(code) => return return_xl_error(code),
    };
    let mut cells: Vec<XLOPER12> = Vec::with_capacity((n_cols + 1) * 2);
    cells.push(XLOPER12::from_str("Column"));
    cells.push(XLOPER12::from_str("Sum"));
    for (i, col) in cols.iter().enumerate() {
        let sum: f64 = col.iter().sum();
        cells.push(XLOPER12::from_str(&format!("X{}", i + 1)));
        cells.push(XLOPER12::from_f64(sum));
    }
    build_multi(cells, n_cols + 1, 2)
}

#[no_mangle]
pub extern "system" fn xl_sci_scale(
    y_range: *const XLOPER12,
    factor_arg: *const XLOPER12,
    center_arg: *const XLOPER12,
) -> *mut XLOPER12 {
    let mut y = match xloper_to_f64_vec(y_range) {
        Ok(v) => v,
        Err(code) => return return_xl_error(code),
    };
    let factor = match parse_optional_f64(factor_arg, 1.0) {
        Ok(v) => v,
        Err(code) => return return_xl_error(code),
    };
    let center = match parse_optional_bool(center_arg, false) {
        Ok(v) => v,
        Err(code) => return return_xl_error(code),
    };
    if center && !y.is_empty() {
        let mean = y.iter().sum::<f64>() / y.len() as f64;
        for v in &mut y {
            *v -= mean;
        }
    }
    for v in &mut y {
        *v *= factor;
    }

    let mut cells: Vec<XLOPER12> = Vec::with_capacity(y.len() + 1);
    cells.push(XLOPER12::from_str("Scaled"));
    for v in y {
        cells.push(XLOPER12::from_f64(v));
    }
    let rows = cells.len();
    build_multi(cells, rows, 1)
}

// #[no_mangle]
// pub extern "system" fn xl_sci_ols(
//     y_range: *const XLOPER12,
//     x_range: *const XLOPER12,
// ) -> *mut XLOPER12 {
//     let y = match xloper_to_f64_vec(y_range) {
//         Ok(v) => v,
//         Err(code) => return return_xl_error(code),
//     };
//     let (x_vars, _n_rows, n_cols) = match xloper_to_columns(x_range) {
//         Ok(v) => v,
//         Err(code) => return return_xl_error(code),
//     };
//     let mut names = vec!["Intercept".to_string()];
//     for i in 1..=n_cols {
//         names.push(format!("X{}", i));
//     }
//     let result = match ols_regression(&y, &x_vars, &names) {
//         Ok(r) => r,
//         Err(_) => return return_xl_error(XLERR_NUM),
//     };

//     // Return a compact 3-column output: Term | Coef | StdErr
//     let n_rows = result.coefficients.len() + 1;
//     let mut cells: Vec<XLOPER12> = Vec::with_capacity(n_rows * 3);
//     cells.push(XLOPER12::from_str("Term"));
//     cells.push(XLOPER12::from_str("Coef"));
//     cells.push(XLOPER12::from_str("StdErr"));
//     for i in 0..result.coefficients.len() {
//         cells.push(XLOPER12::from_str(&result.variable_names[i]));
//         cells.push(XLOPER12::from_f64(result.coefficients[i]));
//         cells.push(XLOPER12::from_f64(result.std_errors[i]));
//     }
//     build_multi(cells, n_rows, 3)
// }


#[no_mangle]
pub extern "system" fn xl_sci_regdiag() -> *mut XLOPER12 {
    let msg = REG_LOG.get().map(|s| s.as_str()).unwrap_or("(not set)");
    XlReturn::str(msg).into_raw()
}

#[no_mangle]
pub extern "system" fn xl_sci_nil() -> *mut XLOPER12 {
    XlReturn::nil().into_raw()
}

#[no_mangle]
pub extern "system" fn xl_sci_blank() -> *mut XLOPER12 {
    XlReturn::str("").into_raw()
}

#[no_mangle]
pub extern "system" fn xl_sci_refinfo(reference: *const XLOPER12) -> *mut XLOPER12 {
    if reference.is_null() {
        return return_xl_error(XLERR_VALUE);
    }
    let oper = unsafe { &*reference };
    let base = oper.base_type();

    let mut areas: Vec<XLREF12> = Vec::new();
    let mut sheet_id: Option<usize> = None;
    match base {
        XLTYPE_SREF => {
            let sref = unsafe { &*std::ptr::addr_of!(oper.val.sref) };
            areas.push(sref.ref_);
        }
        XLTYPE_REF => {
            let mref = unsafe { &*std::ptr::addr_of!(oper.val.mref) };
            sheet_id = Some(mref.id_sheet);
            if mref.lpmref.is_null() {
                return return_xl_error(XLERR_VALUE);
            }
            let tbl = unsafe { &*mref.lpmref };
            let count = tbl.count as usize;
            let base_ptr = tbl.reftbl.as_ptr();
            for i in 0..count {
                let area = unsafe { *base_ptr.add(i) };
                areas.push(area);
            }
        }
        _ => return return_xl_error(XLERR_VALUE),
    }

    let sheet_name = sheet_name_from_ref(reference).unwrap_or_default();

    let mut cells: Vec<XLOPER12> = Vec::with_capacity((4 + areas.len()) * 2);
    cells.push(XLOPER12::from_str("Type"));
    cells.push(XLOPER12::from_str(if base == XLTYPE_SREF { "SREF" } else { "REF" }));
    cells.push(XLOPER12::from_str("Sheet"));
    cells.push(XLOPER12::from_str(&sheet_name));
    if let Some(id) = sheet_id {
        cells.push(XLOPER12::from_str("SheetId"));
        cells.push(XLOPER12::from_f64(id as f64));
    }
    cells.push(XLOPER12::from_str("Areas"));
    cells.push(XLOPER12::from_f64(areas.len() as f64));
    for (i, area) in areas.iter().enumerate() {
        let label = format!("Area {}", i + 1);
        let range = format!(
            "R{}C{}:R{}C{}",
            area.rw_first + 1,
            area.col_first + 1,
            area.rw_last + 1,
            area.col_last + 1
        );
        cells.push(XLOPER12::from_str(&label));
        cells.push(XLOPER12::from_str(&range));
    }
    let rows = cells.len() / 2;
    build_multi(cells, rows, 2)
}

#[no_mangle]
pub extern "system" fn xl_sci_refvalues(reference: *const XLOPER12) -> *mut XLOPER12 {
    if reference.is_null() {
        return return_xl_error(XLERR_VALUE);
    }
    let oper = unsafe { &*reference };
    match coerce_to_owned(oper) {
        Ok(p) => p,
        Err(code) => return_xl_error(code),
    }
}

#[no_mangle]
pub extern "system" fn xl_sci_flow() -> *mut XLOPER12 {
    let oper = XLOPER12 {
        val: XLOPER12Val {
            flow: ManuallyDrop::new(XLOPER12Flow {
                valflow: XLOPER12FlowVal { level: 1 },
                rw: 0,
                col: 0,
                xlflow: XLFLOW_RESTART,
            }),
        },
        xltype: XLTYPE_FLOW,
    };
    XlReturn::from_oper(oper).into_raw()
}

#[no_mangle]
pub extern "system" fn xl_sci_bigdata() -> *mut XLOPER12 {
    let (ret, mut res) = excel12(XL_GET_INST_PTR, &mut []);
    if ret != XLRET_SUCCESS {
        return return_xl_error(XLERR_VALUE);
    }
    if res.base_type() != XLTYPE_BIGDATA {
        if (res.xltype & XLBIT_XL_FREE) != 0 {
            excel_free(&mut res);
        }
        return return_xl_error(XLERR_VALUE);
    }

    let big = unsafe { &*std::ptr::addr_of!(res.val.bigdata) };
    let handle = unsafe { big.h.hdata };
    let cb = big.cb_data;
    let oper = XLOPER12 {
        val: XLOPER12Val {
            bigdata: ManuallyDrop::new(XLOPER12BigData {
                h: XLOPER12BigDataHandle { hdata: handle },
                cb_data: cb,
            }),
        },
        xltype: XLTYPE_BIGDATA,
    };
    if (res.xltype & XLBIT_XL_FREE) != 0 {
        excel_free(&mut res);
    }
    XlReturn::from_oper(oper).into_raw()
}

#[no_mangle]
pub extern "system" fn xl_sci_error(code: *const XLOPER12) -> *mut XLOPER12 {
    if code.is_null() {
        return return_xl_error(XLERR_VALUE);
    }
    let oper = unsafe { &*code };
    let err = match oper.base_type() {
        XLTYPE_NUM => (unsafe { oper.val.num }) as i32,
        XLTYPE_INT => unsafe { oper.val.w },
        XLTYPE_ERR => unsafe { oper.val.err },
        _ => return return_xl_error(XLERR_VALUE),
    };
    return_xl_error(err)
}

#[no_mangle]
pub extern "system" fn xl_sci_toint(value: *const XLOPER12) -> *mut XLOPER12 {
    if value.is_null() {
        return return_xl_error(XLERR_VALUE);
    }
    let oper = unsafe { &*value };
    let v = match oper.base_type() {
        XLTYPE_NUM => unsafe { oper.val.num },
        XLTYPE_INT => (unsafe { oper.val.w }) as f64,
        XLTYPE_ERR => return return_xl_error(unsafe { oper.val.err }),
        _ => return return_xl_error(XLERR_VALUE),
    };
    XlReturn::int(v as i32).into_raw()
}

#[no_mangle]
pub extern "system" fn xl_sci_threshold(
    value: *const XLOPER12,
    cutoff_arg: *const XLOPER12,
    strict_arg: *const XLOPER12,
) -> *mut XLOPER12 {
    if value.is_null() {
        return return_xl_error(XLERR_VALUE);
    }
    let oper = unsafe { &*value };
    let v = match oper.base_type() {
        XLTYPE_NUM => unsafe { oper.val.num },
        XLTYPE_INT => (unsafe { oper.val.w }) as f64,
        XLTYPE_ERR => return return_xl_error(unsafe { oper.val.err }),
        _ => return return_xl_error(XLERR_VALUE),
    };
    let cutoff = match parse_optional_f64(cutoff_arg, 0.0) {
        Ok(v) => v,
        Err(code) => return return_xl_error(code),
    };
    let strict = match parse_optional_bool(strict_arg, false) {
        Ok(v) => v,
        Err(code) => return return_xl_error(code),
    };
    let pass = if strict { v > cutoff } else { v >= cutoff };
    XlReturn::bool(pass).into_raw()
}

#[no_mangle]
pub extern "system" fn xl_sci_types() -> *mut XLOPER12 {
    let mut cells: Vec<XLOPER12> = Vec::with_capacity(30);
    // Header
    cells.push(XLOPER12::from_str("Type"));
    cells.push(XLOPER12::from_str("Value"));
    cells.push(XLOPER12::from_str("xltype"));
    // xltypeStr
    cells.push(XLOPER12::from_str("String"));
    cells.push(XLOPER12::from_str("alpha"));
    cells.push(XLOPER12::from_str("0x0002"));
    // xltypeNum
    cells.push(XLOPER12::from_str("Num (f64)"));
    cells.push(XLOPER12::from_f64(3.14159));
    cells.push(XLOPER12::from_str("0x0001"));
    // xltypeInt
    cells.push(XLOPER12::from_str("Int (i32)"));
    cells.push(XLOPER12::from_int(42));
    cells.push(XLOPER12::from_str("0x0800"));
    // xltypeBool TRUE
    cells.push(XLOPER12::from_str("Bool (TRUE)"));
    cells.push(XLOPER12::from_bool(true));
    cells.push(XLOPER12::from_str("0x0004"));
    // xltypeBool FALSE
    cells.push(XLOPER12::from_str("Bool (FALSE)"));
    cells.push(XLOPER12::from_bool(false));
    cells.push(XLOPER12::from_str("0x0004"));
    // xltypeErr variants
    cells.push(XLOPER12::from_str("Err #NULL!"));
    cells.push(XLOPER12::from_err(XLERR_NULL));
    cells.push(XLOPER12::from_str("0x0010"));
    cells.push(XLOPER12::from_str("Err #DIV/0!"));
    cells.push(XLOPER12::from_err(XLERR_DIV0));
    cells.push(XLOPER12::from_str("0x0010"));
    cells.push(XLOPER12::from_str("Err #VALUE!"));
    cells.push(XLOPER12::from_err(XLERR_VALUE));
    cells.push(XLOPER12::from_str("0x0010"));
    cells.push(XLOPER12::from_str("Err #REF!"));
    cells.push(XLOPER12::from_err(XLERR_REF));
    cells.push(XLOPER12::from_str("0x0010"));
    cells.push(XLOPER12::from_str("Err #NAME?"));
    cells.push(XLOPER12::from_err(XLERR_NAME));
    cells.push(XLOPER12::from_str("0x0010"));
    cells.push(XLOPER12::from_str("Err #NUM!"));
    cells.push(XLOPER12::from_err(XLERR_NUM));
    cells.push(XLOPER12::from_str("0x0010"));
    cells.push(XLOPER12::from_str("Err #N/A"));
    cells.push(XLOPER12::from_err(XLERR_NA));
    cells.push(XLOPER12::from_str("0x0010"));
    // xltypeNil
    cells.push(XLOPER12::from_str("Nil"));
    cells.push(XLOPER12::nil());
    cells.push(XLOPER12::from_str("0x0100"));
    // xltypeMissing
    cells.push(XLOPER12::from_str("Missing"));
    cells.push(XLOPER12::missing());
    cells.push(XLOPER12::from_str("0x0080"));

    let n_rows = cells.len() / 3;
    build_multi(cells, n_rows, 3)
}

#[no_mangle]
pub extern "system" fn xl_sci_dims(x_range: *const XLOPER12) -> *mut XLOPER12 {
    let (_cols, n_rows, n_cols) = match xloper_to_columns(x_range) {
        Ok(v) => v,
        Err(code) => return return_xl_error(code),
    };
    build_kv_table(&[("Rows", n_rows as f64), ("Cols", n_cols as f64)])
}

// ── Excel callbacks ─────────────────────────────────────────────────────────
#[no_mangle]
pub extern "system" fn xlAutoOpen() -> i32 {
    let reg = Reg::new();

    reg.add("xl_sci_version", "Q$", "SCI.VERSION", "", "xll-rs", "Library version", &[]);
    reg.add("xl_sci_hello", "Q$", "SCI.HELLO", "", "xll-rs", "Greeting string", &[]);
    reg.add("xl_sci_add", "QQQ$", "SCI.ADD", "a, b", "xll-rs", "Add two numbers", &["A", "B"]);
    reg.add(
        "xl_sci_not",
        "QQ$",
        "SCI.NOT",
        "value",
        "xll-rs",
        "Logical NOT",
        &["Boolean value"],
    );
    reg.add(
        "xl_sci_echo",
        "QQ$",
        "SCI.ECHO",
        "text",
        "xll-rs",
        "Echo a string",
        &["Input text"],
    );
    reg.add(
        "xl_sci_mean",
        "QQ$",
        "SCI.MEAN",
        "y_range",
        "xll-rs",
        "Mean of a numeric range",
        &["Numeric range"],
    );
    reg.add(
        "xl_sci_describe",
        "QQ$",
        "SCI.DESCRIBE",
        "y_range",
        "xll-rs",
        "Summary stats (mean, std, min, max)",
        &["Numeric range"],
    );
    reg.add(
        "xl_sci_corr",
        "QQQ$",
        "SCI.CORR",
        "y_range, x_range",
        "xll-rs",
        "Correlation between two ranges",
        &["Y range", "X range"],
    );
    reg.add(
        "xl_sci_colsum",
        "QQ$",
        "SCI.COLSUM",
        "x_range",
        "xll-rs",
        "Column sums of a matrix",
        &["Matrix range"],
    );
    reg.add(
        "xl_sci_scale",
        "QQQQ$",
        "SCI.SCALE",
        "y_range, [factor], [center]",
        "xll-rs",
        "Scale a vector with optional centering",
        &["Numeric range", "Scale factor (default 1)", "Center? (default FALSE)"],
    );
    // reg.add(
    //     "xl_sci_ols",
    //     "QQQ$",
    //     "SCI.OLS",
    //     "y_range, x_range",
    //     "xll-rs",
    //     "OLS regression (compact output)",
    //     &["Response range", "Predictor matrix"],
    // );
    reg.add("xl_sci_nil", "Q$", "SCI.NIL", "", "xll-rs", "Return xltypeNil", &[]);
    reg.add(
        "xl_sci_blank",
        "Q$",
        "SCI.BLANK",
        "",
        "xll-rs",
        "Return empty string (blank cell)",
        &[],
    );
    reg.add(
        "xl_sci_error",
        "QQ$",
        "SCI.ERROR",
        "code",
        "xll-rs",
        "Return an Excel error by code",
        &["Excel error code"],
    );
    reg.add(
        "xl_sci_toint",
        "QQ$",
        "SCI.TOINT",
        "value",
        "xll-rs",
        "Truncate a numeric value to int",
        &["Numeric value"],
    );
    reg.add(
        "xl_sci_threshold",
        "QQQQ$",
        "SCI.THRESH",
        "value, [cutoff], [strict]",
        "xll-rs",
        "Threshold comparison with optional args",
        &["Value", "Cutoff (default 0)", "Strict? (default FALSE)"],
    );
    reg.add(
        "xl_sci_types",
        "Q$",
        "SCI.TYPES",
        "",
        "xll-rs",
        "Mixed-type table",
        &[],
    );
    reg.add(
        "xl_sci_dims",
        "QQ$",
        "SCI.DIMS",
        "x_range",
        "xll-rs",
        "Rows and columns for a numeric range",
        &["Numeric matrix"],
    );
    let r1 = reg.add(
        "xl_sci_refinfo",
        "QU#",
        "SCI.REFINFO",
        "ref",
        "xll-rs",
        "Reference metadata (SREF/REF)",
        &["Reference or range"],
    );
    let r2 = reg.add(
        "xl_sci_refvalues",
        "QU#",
        "SCI.REFVALUES",
        "ref",
        "xll-rs",
        "Coerce reference to values",
        &["Reference or range"],
    );

    let r3 = reg.add("xl_sci_regdiag", "Q$", "SCI.REGDIAG", "", "xll-rs", "Registration diagnostics", &[]);

    let _ = REG_LOG.set(format!(
        "REFINFO(QU#)={} REFVALUES(QU#)={} REGDIAG(Q$)={}",
        r1, r2, r3
    ));
    reg.add(
        "xl_sci_flow",
        "Q",
        "SCI.FLOW",
        "",
        "xll-rs",
        "Return xltypeFlow (testing)",
        &[],
    );
    reg.add(
        "xl_sci_bigdata",
        "Q",
        "SCI.BIGDATA",
        "",
        "xll-rs",
        "Return xltypeBigData via xlGetInstPtr",
        &[],
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
            return XlReturn::str("xll-rs scientific example").into_raw();
        }
    }
    return_xl_error(XLERR_VALUE)
}

// Excel calls this after it copies results with xlbitDLLFree set
pub use xll_rs::memory::xlAutoFree12;
