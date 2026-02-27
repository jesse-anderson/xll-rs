//! Conversion helpers between Excel XLOPER12 ranges and Rust types.
//!
//! # Gotchas discovered during development
//!
//! - When arguments are registered as type `Q`, Excel coerces cell references
//!   into `xltypeMulti` arrays automatically.  We never see `xltypeSRef` or
//!   `xltypeRef` — those only appear with type `U` registration.
//! - A single cell passed as an argument arrives as `xltypeNum`, `xltypeStr`,
//!   etc. — NOT as a 1x1 `xltypeMulti`.  Must handle both cases.
//! - Empty cells inside a multi-cell range arrive as `xltypeNil`.
//! - Cells containing errors arrive as `xltypeErr` — we propagate them.

use super::entrypoint::{excel12, excel_free, XLRET_SUCCESS};
use super::returning::XlReturn;
use super::types::*;

/// Extract a column vector of f64 from an XLOPER12 (single value or range).
///
/// For a 2D range, reads all values in column-major order flattened into one
/// vector (used for the Y variable — must be a single column).
pub fn xloper_to_f64_vec(p: *const XLOPER12) -> Result<Vec<f64>, i32> {
    if p.is_null() {
        return Err(XLERR_VALUE);
    }
    let oper = unsafe { &*p };
    match oper.base_type() {
        XLTYPE_NUM => Ok(vec![unsafe { oper.val.num }]),
        XLTYPE_INT => Ok(vec![unsafe { oper.val.w } as f64]),
        XLTYPE_MISSING | XLTYPE_NIL => Err(XLERR_VALUE),
        XLTYPE_ERR => Err(unsafe { oper.val.err }),
        XLTYPE_MULTI => {
            let arr = unsafe { &*std::ptr::addr_of!(oper.val.array) };
            if arr.rows <= 0 || arr.columns <= 0 {
                return Err(XLERR_VALUE);
            }
            let total = (arr.rows * arr.columns) as usize;
            let mut result = Vec::with_capacity(total);
            for i in 0..total {
                let elem = unsafe { &*arr.lparray.add(i) };
                match elem.base_type() {
                    XLTYPE_NUM => result.push(unsafe { elem.val.num }),
                    XLTYPE_INT => result.push(unsafe { elem.val.w } as f64),
                    XLTYPE_NIL => return Err(XLERR_VALUE), // empty cell in data
                    XLTYPE_ERR => return Err(unsafe { elem.val.err }),
                    _ => return Err(XLERR_VALUE),
                }
            }
            Ok(result)
        }
        _ => Err(XLERR_VALUE),
    }
}

/// Extract a 2D range into column-major `Vec<Vec<f64>>` — the format
/// `ols_regression` expects for `x_vars`.
///
/// A range of shape (n_rows, n_cols) becomes `n_cols` vectors each of length
/// `n_rows`.  Also returns `(n_rows, n_cols)`.
pub fn xloper_to_columns(p: *const XLOPER12) -> Result<(Vec<Vec<f64>>, usize, usize), i32> {
    if p.is_null() {
        return Err(XLERR_VALUE);
    }
    let oper = unsafe { &*p };
    match oper.base_type() {
        // Single numeric value → 1 column, 1 row
        XLTYPE_NUM => Ok((vec![vec![unsafe { oper.val.num }]], 1, 1)),
        XLTYPE_INT => Ok((vec![vec![unsafe { oper.val.w } as f64]], 1, 1)),
        XLTYPE_MISSING | XLTYPE_NIL => Err(XLERR_VALUE),
        XLTYPE_ERR => Err(unsafe { oper.val.err }),
        XLTYPE_MULTI => {
            let arr = unsafe { &*std::ptr::addr_of!(oper.val.array) };
            if arr.rows <= 0 || arr.columns <= 0 {
                return Err(XLERR_VALUE);
            }
            let n_rows = arr.rows as usize;
            let n_cols = arr.columns as usize;

            // Build column-major: columns[col][row]
            let mut columns = vec![Vec::with_capacity(n_rows); n_cols];
            for row in 0..n_rows {
                for col in 0..n_cols {
                    let idx = row * n_cols + col; // row-major in XLOPER12
                    let elem = unsafe { &*arr.lparray.add(idx) };
                    match elem.base_type() {
                        XLTYPE_NUM => columns[col].push(unsafe { elem.val.num }),
                        XLTYPE_INT => columns[col].push(unsafe { elem.val.w } as f64),
                        XLTYPE_NIL => return Err(XLERR_VALUE),
                        XLTYPE_ERR => return Err(unsafe { elem.val.err }),
                        _ => return Err(XLERR_VALUE),
                    }
                }
            }
            Ok((columns, n_rows, n_cols))
        }
        _ => Err(XLERR_VALUE),
    }
}

/// Extract a column vector of i32 from an XLOPER12 (single value or range).
pub fn xloper_to_i32_vec(p: *const XLOPER12) -> Result<Vec<i32>, i32> {
    if p.is_null() {
        return Err(XLERR_VALUE);
    }
    let oper = unsafe { &*p };
    match oper.base_type() {
        XLTYPE_INT => Ok(vec![unsafe { oper.val.w }]),
        XLTYPE_NUM => Ok(vec![unsafe { oper.val.num } as i32]),
        XLTYPE_MISSING | XLTYPE_NIL => Err(XLERR_VALUE),
        XLTYPE_ERR => Err(unsafe { oper.val.err }),
        XLTYPE_MULTI => {
            let arr = unsafe { &*std::ptr::addr_of!(oper.val.array) };
            if arr.rows <= 0 || arr.columns <= 0 {
                return Err(XLERR_VALUE);
            }
            let total = (arr.rows * arr.columns) as usize;
            let mut result = Vec::with_capacity(total);
            for i in 0..total {
                let elem = unsafe { &*arr.lparray.add(i) };
                match elem.base_type() {
                    XLTYPE_INT => result.push(unsafe { elem.val.w }),
                    XLTYPE_NUM => result.push(unsafe { elem.val.num } as i32),
                    XLTYPE_NIL => return Err(XLERR_VALUE),
                    XLTYPE_ERR => return Err(unsafe { elem.val.err }),
                    _ => return Err(XLERR_VALUE),
                }
            }
            Ok(result)
        }
        _ => Err(XLERR_VALUE),
    }
}

/// Extract a column vector of bool from an XLOPER12 (single value or range).
pub fn xloper_to_bool_vec(p: *const XLOPER12) -> Result<Vec<bool>, i32> {
    if p.is_null() {
        return Err(XLERR_VALUE);
    }
    let oper = unsafe { &*p };
    match oper.base_type() {
        XLTYPE_BOOL => Ok(vec![unsafe { oper.val.xbool } != 0]),
        XLTYPE_INT => Ok(vec![unsafe { oper.val.w } != 0]),
        XLTYPE_NUM => Ok(vec![unsafe { oper.val.num } != 0.0]),
        XLTYPE_MISSING | XLTYPE_NIL => Err(XLERR_VALUE),
        XLTYPE_ERR => Err(unsafe { oper.val.err }),
        XLTYPE_MULTI => {
            let arr = unsafe { &*std::ptr::addr_of!(oper.val.array) };
            if arr.rows <= 0 || arr.columns <= 0 {
                return Err(XLERR_VALUE);
            }
            let total = (arr.rows * arr.columns) as usize;
            let mut result = Vec::with_capacity(total);
            for i in 0..total {
                let elem = unsafe { &*arr.lparray.add(i) };
                match elem.base_type() {
                    XLTYPE_BOOL => result.push(unsafe { elem.val.xbool } != 0),
                    XLTYPE_INT => result.push(unsafe { elem.val.w } != 0),
                    XLTYPE_NUM => result.push(unsafe { elem.val.num } != 0.0),
                    XLTYPE_NIL => return Err(XLERR_VALUE),
                    XLTYPE_ERR => return Err(unsafe { elem.val.err }),
                    _ => return Err(XLERR_VALUE),
                }
            }
            Ok(result)
        }
        _ => Err(XLERR_VALUE),
    }
}

/// Extract a row-major matrix of f64 from an XLOPER12.
pub fn xloper_to_rows_f64(p: *const XLOPER12) -> Result<Vec<Vec<f64>>, i32> {
    if p.is_null() {
        return Err(XLERR_VALUE);
    }
    let oper = unsafe { &*p };
    match oper.base_type() {
        XLTYPE_NUM => Ok(vec![vec![unsafe { oper.val.num }]]),
        XLTYPE_INT => Ok(vec![vec![unsafe { oper.val.w } as f64]]),
        XLTYPE_MISSING | XLTYPE_NIL => Err(XLERR_VALUE),
        XLTYPE_ERR => Err(unsafe { oper.val.err }),
        XLTYPE_MULTI => {
            let arr = unsafe { &*std::ptr::addr_of!(oper.val.array) };
            if arr.rows <= 0 || arr.columns <= 0 {
                return Err(XLERR_VALUE);
            }
            let n_rows = arr.rows as usize;
            let n_cols = arr.columns as usize;
            let mut rows = vec![Vec::with_capacity(n_cols); n_rows];
            for row in 0..n_rows {
                for col in 0..n_cols {
                    let idx = row * n_cols + col;
                    let elem = unsafe { &*arr.lparray.add(idx) };
                    match elem.base_type() {
                        XLTYPE_NUM => rows[row].push(unsafe { elem.val.num }),
                        XLTYPE_INT => rows[row].push(unsafe { elem.val.w } as f64),
                        XLTYPE_NIL => return Err(XLERR_VALUE),
                        XLTYPE_ERR => return Err(unsafe { elem.val.err }),
                        _ => return Err(XLERR_VALUE),
                    }
                }
            }
            Ok(rows)
        }
        _ => Err(XLERR_VALUE),
    }
}

/// Extract a row-major matrix of i32 from an XLOPER12.
pub fn xloper_to_rows_i32(p: *const XLOPER12) -> Result<Vec<Vec<i32>>, i32> {
    if p.is_null() {
        return Err(XLERR_VALUE);
    }
    let oper = unsafe { &*p };
    match oper.base_type() {
        XLTYPE_INT => Ok(vec![vec![unsafe { oper.val.w }]]),
        XLTYPE_NUM => Ok(vec![vec![unsafe { oper.val.num } as i32]]),
        XLTYPE_MISSING | XLTYPE_NIL => Err(XLERR_VALUE),
        XLTYPE_ERR => Err(unsafe { oper.val.err }),
        XLTYPE_MULTI => {
            let arr = unsafe { &*std::ptr::addr_of!(oper.val.array) };
            if arr.rows <= 0 || arr.columns <= 0 {
                return Err(XLERR_VALUE);
            }
            let n_rows = arr.rows as usize;
            let n_cols = arr.columns as usize;
            let mut rows = vec![Vec::with_capacity(n_cols); n_rows];
            for row in 0..n_rows {
                for col in 0..n_cols {
                    let idx = row * n_cols + col;
                    let elem = unsafe { &*arr.lparray.add(idx) };
                    match elem.base_type() {
                        XLTYPE_INT => rows[row].push(unsafe { elem.val.w }),
                        XLTYPE_NUM => rows[row].push(unsafe { elem.val.num } as i32),
                        XLTYPE_NIL => return Err(XLERR_VALUE),
                        XLTYPE_ERR => return Err(unsafe { elem.val.err }),
                        _ => return Err(XLERR_VALUE),
                    }
                }
            }
            Ok(rows)
        }
        _ => Err(XLERR_VALUE),
    }
}

/// Extract a row-major matrix of bool from an XLOPER12.
pub fn xloper_to_rows_bool(p: *const XLOPER12) -> Result<Vec<Vec<bool>>, i32> {
    if p.is_null() {
        return Err(XLERR_VALUE);
    }
    let oper = unsafe { &*p };
    match oper.base_type() {
        XLTYPE_BOOL => Ok(vec![vec![unsafe { oper.val.xbool } != 0]]),
        XLTYPE_INT => Ok(vec![vec![unsafe { oper.val.w } != 0]]),
        XLTYPE_NUM => Ok(vec![vec![unsafe { oper.val.num } != 0.0]]),
        XLTYPE_MISSING | XLTYPE_NIL => Err(XLERR_VALUE),
        XLTYPE_ERR => Err(unsafe { oper.val.err }),
        XLTYPE_MULTI => {
            let arr = unsafe { &*std::ptr::addr_of!(oper.val.array) };
            let n_rows = arr.rows as usize;
            let n_cols = arr.columns as usize;
            if n_rows == 0 || n_cols == 0 {
                return Err(XLERR_VALUE);
            }
            let mut rows = vec![Vec::with_capacity(n_cols); n_rows];
            for row in 0..n_rows {
                for col in 0..n_cols {
                    let idx = row * n_cols + col;
                    let elem = unsafe { &*arr.lparray.add(idx) };
                    match elem.base_type() {
                        XLTYPE_BOOL => rows[row].push(unsafe { elem.val.xbool } != 0),
                        XLTYPE_INT => rows[row].push(unsafe { elem.val.w } != 0),
                        XLTYPE_NUM => rows[row].push(unsafe { elem.val.num } != 0.0),
                        XLTYPE_NIL => return Err(XLERR_VALUE),
                        XLTYPE_ERR => return Err(unsafe { elem.val.err }),
                        _ => return Err(XLERR_VALUE),
                    }
                }
            }
            Ok(rows)
        }
        _ => Err(XLERR_VALUE),
    }
}

/// Parse an optional numeric argument.
///
/// Returns `default` if the argument is `xltypeMissing` or `xltypeNil`.
pub fn parse_optional_f64(p: *const XLOPER12, default: f64) -> Result<f64, i32> {
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

/// Parse an optional boolean argument.
///
/// Returns `default` if the argument is `xltypeMissing` or `xltypeNil`.
pub fn parse_optional_bool(p: *const XLOPER12, default: bool) -> Result<bool, i32> {
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

/// Build an xltypeMulti XLOPER12 from a grid of cells.
///
/// `cells` is a flat row-major array of XLOPER12 values.  The returned
/// XLOPER12 is heap-allocated with `xlbitDLLFree` set — caller returns it
/// directly to Excel, which will call `xlAutoFree12` after copying.
///
/// # Memory contract
///
/// - The `cells` Vec is leaked (forgotten) and reclaimed in `xlAutoFree12`.
/// - Any string XLOPER12s inside `cells` must have been created with
///   `XLOPER12::from_str()` (which allocates its own buffer).  They are
///   freed recursively in `xlAutoFree12`.
/// - Numeric / error / nil cells have no allocations.
pub fn build_multi(mut cells: Vec<XLOPER12>, rows: usize, cols: usize) -> *mut XLOPER12 {
    debug_assert_eq!(cells.len(), rows * cols);

    let lparray = cells.as_mut_ptr();
    std::mem::forget(cells);

    let result = Box::new(XLOPER12 {
        val: XLOPER12Val {
            array: std::mem::ManuallyDrop::new(XLOPER12Array {
                lparray,
                rows: rows as i32,
                columns: cols as i32,
            }),
        },
        xltype: XLTYPE_MULTI | XLBIT_DLL_FREE,
    });
    Box::into_raw(result)
}

/// Return a heap-allocated error XLOPER12 suitable as a UDF return value.
pub fn return_xl_error(code: i32) -> *mut XLOPER12 {
    Box::into_raw(Box::new(XLOPER12 {
        val: XLOPER12Val { err: code },
        xltype: XLTYPE_ERR | XLBIT_DLL_FREE,
    }))
}

/// Convert column-major vectors into a row-major flat buffer.
///
/// Input is `columns[col][row]`, output is flattened row-by-row.
pub fn columns_to_row_major(columns: &[Vec<f64>]) -> Result<Vec<f64>, i32> {
    if columns.is_empty() {
        return Ok(Vec::new());
    }
    let n_rows = columns[0].len();
    if !columns.iter().all(|c| c.len() == n_rows) {
        return Err(XLERR_VALUE);
    }
    let n_cols = columns.len();
    let mut data = Vec::with_capacity(n_rows * n_cols);
    for row in 0..n_rows {
        for col in 0..n_cols {
            data.push(columns[col][row]);
        }
    }
    Ok(data)
}

/// Convert row-major vectors into a row-major flat buffer.
///
/// Input is `rows[row][col]`, output is flattened row-by-row.
pub fn rows_to_row_major(rows: &[Vec<f64>]) -> Result<Vec<f64>, i32> {
    if rows.is_empty() {
        return Ok(Vec::new());
    }
    let n_cols = rows[0].len();
    if !rows.iter().all(|r| r.len() == n_cols) {
        return Err(XLERR_VALUE);
    }
    let mut data = Vec::with_capacity(rows.len() * n_cols);
    for row in rows {
        data.extend_from_slice(row);
    }
    Ok(data)
}

/// Build a 2-column label/value table for returning to Excel.
///
/// Each input row becomes: `[label, value]`.
pub fn build_kv_table(rows: &[(&str, f64)]) -> *mut XLOPER12 {
    let n_rows = rows.len();
    let mut cells: Vec<XLOPER12> = Vec::with_capacity(n_rows * 2);
    for (label, value) in rows {
        cells.push(XLOPER12::from_str(label));
        cells.push(XLOPER12::from_f64(*value));
    }
    build_multi(cells, n_rows, 2)
}

/// Resolve a sheet name from a reference using `xlSheetNm`.
///
/// If the input is already a string (`xltypeStr`), it is returned directly.
pub fn sheet_name_from_ref(p: *const XLOPER12) -> Option<String> {
    if p.is_null() {
        return None;
    }
    let oper = unsafe { &*p };
    if oper.base_type() == XLTYPE_STR {
        return oper.as_string();
    }
    match oper.base_type() {
        XLTYPE_SREF | XLTYPE_REF => {}
        _ => return None,
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

/// Deep-clone an Excel-owned XLOPER12 into DLL-owned memory.
///
/// Supported types: num, int, bool, err, nil, str, multi.
pub fn clone_excel_oper(oper: &XLOPER12) -> Result<*mut XLOPER12, i32> {
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
            if arr.rows <= 0 || arr.columns <= 0 {
                return Err(XLERR_VALUE);
            }
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

/// Coerce a reference (SREF/REF) to values and return a DLL-owned multi.
pub fn coerce_to_owned(oper: &XLOPER12) -> Result<*mut XLOPER12, i32> {
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
