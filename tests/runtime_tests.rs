use std::mem::ManuallyDrop;

use xll_rs::convert::{
    coerce_to_owned, columns_to_row_major, parse_optional_bool, parse_optional_f64,
    rows_to_row_major, sheet_name_from_ref, xloper_to_columns, xloper_to_f64_vec,
};
use xll_rs::entrypoint::init_test_entrypoint;
use xll_rs::memory::free_xloper_recursive;
use xll_rs::types::*;

fn make_multi(cells: &mut Vec<XLOPER12>, rows: i32, cols: i32) -> XLOPER12 {
    XLOPER12 {
        val: XLOPER12Val {
            array: ManuallyDrop::new(XLOPER12Array {
                lparray: cells.as_mut_ptr(),
                rows,
                columns: cols,
            }),
        },
        xltype: XLTYPE_MULTI,
    }
}

#[test]
fn base_type_masks_flags() {
    let oper = XLOPER12::from_str("mask");
    assert_eq!(oper.base_type(), XLTYPE_STR);
    assert_ne!(oper.xltype & XLBIT_DLL_FREE, 0);
}

#[test]
fn missing_nil_err_roundtrip() {
    let missing = XLOPER12::missing();
    assert!(missing.is_missing());
    assert_eq!(missing.base_type(), XLTYPE_MISSING);

    let nil = XLOPER12::nil();
    assert_eq!(nil.base_type(), XLTYPE_NIL);

    let err = XLOPER12::from_err(XLERR_NUM);
    assert_eq!(err.base_type(), XLTYPE_ERR);
    unsafe {
        assert_eq!(err.val.err, XLERR_NUM);
    }
}

#[test]
fn from_str_rejects_too_long() {
    let s = "a".repeat(32768);
    let oper = XLOPER12::from_str(&s);
    assert_eq!(oper.base_type(), XLTYPE_ERR);
    unsafe {
        assert_eq!(oper.val.err, XLERR_VALUE);
    }
}

#[test]
fn flow_and_bigdata_types() {
    let flow = XLOPER12 {
        val: XLOPER12Val {
            flow: ManuallyDrop::new(XLOPER12Flow {
                valflow: XLOPER12FlowVal { level: 3 },
                rw: 10,
                col: 2,
                xlflow: XLFLOW_RESTART,
            }),
        },
        xltype: XLTYPE_FLOW,
    };
    assert_eq!(flow.base_type(), XLTYPE_FLOW);

    let bigdata = XLOPER12 {
        val: XLOPER12Val {
            bigdata: ManuallyDrop::new(XLOPER12BigData {
                h: XLOPER12BigDataHandle {
                    lpb_data: std::ptr::null_mut(),
                },
                cb_data: 128,
            }),
        },
        xltype: XLTYPE_BIGDATA,
    };
    assert_eq!(bigdata.base_type(), XLTYPE_BIGDATA);
}

#[test]
fn ref_types_roundtrip() {
    let sref = XLOPER12 {
        val: XLOPER12Val {
            sref: ManuallyDrop::new(XLOPER12SRef {
                count: 1,
                ref_: XLREF12 {
                    rw_first: 1,
                    rw_last: 2,
                    col_first: 3,
                    col_last: 4,
                },
            }),
        },
        xltype: XLTYPE_SREF,
    };
    assert_eq!(sref.base_type(), XLTYPE_SREF);

    let mref = XLOPER12MRef {
        lpmref: Box::into_raw(Box::new(XLMREF12 {
            count: 1,
            reftbl: [XLREF12 {
                rw_first: 1,
                rw_last: 1,
                col_first: 1,
                col_last: 1,
            }],
        })),
        id_sheet: 0,
    };
    let mref_oper = XLOPER12 {
        val: XLOPER12Val {
            mref: ManuallyDrop::new(mref),
        },
        xltype: XLTYPE_REF,
    };
    assert_eq!(mref_oper.base_type(), XLTYPE_REF);

    unsafe {
        let ptr = mref_oper.val.mref.lpmref;
        let _ = Box::from_raw(ptr);
    }
}

#[test]
fn free_xloper_recursive_string() {
    let p = Box::into_raw(Box::new(XLOPER12::from_str("free me")));
    unsafe { free_xloper_recursive(p) };
}

#[test]
fn free_xloper_recursive_multi_with_strings() {
    let cells = vec![
        XLOPER12::from_str("A"),
        XLOPER12::from_f64(1.0),
        XLOPER12::from_str("B"),
        XLOPER12::from_f64(2.0),
    ];
    let p = xll_rs::convert::build_multi(cells, 2, 2);
    xll_rs::memory::xlAutoFree12(p);
}

#[test]
fn f64_vec_rejects_non_numeric() {
    let p = Box::into_raw(Box::new(XLOPER12::from_bool(true)));
    let err = xloper_to_f64_vec(p).unwrap_err();
    assert_eq!(err, XLERR_VALUE);
    unsafe { free_xloper_recursive(p) };

    let p = Box::into_raw(Box::new(XLOPER12::from_err(XLERR_REF)));
    let err = xloper_to_f64_vec(p).unwrap_err();
    assert_eq!(err, XLERR_REF);
    unsafe { free_xloper_recursive(p) };

    let p = Box::into_raw(Box::new(XLOPER12::missing()));
    let err = xloper_to_f64_vec(p).unwrap_err();
    assert_eq!(err, XLERR_VALUE);
    unsafe { free_xloper_recursive(p) };
}

#[test]
fn f64_vec_multi_err_propagates() {
    let mut cells = vec![XLOPER12::from_f64(1.0), XLOPER12::from_err(XLERR_NA)];
    let oper = make_multi(&mut cells, 1, 2);
    let err = xloper_to_f64_vec(&oper).unwrap_err();
    assert_eq!(err, XLERR_NA);
}

#[test]
fn columns_reject_invalid_shapes_and_types() {
    let mut empty: Vec<XLOPER12> = Vec::new();
    let oper = make_multi(&mut empty, 0, 0);
    let err = xloper_to_columns(&oper).unwrap_err();
    assert_eq!(err, XLERR_VALUE);

    let p = Box::into_raw(Box::new(XLOPER12::from_str("bad")));
    let err = xloper_to_columns(p).unwrap_err();
    assert_eq!(err, XLERR_VALUE);
    unsafe { free_xloper_recursive(p) };
}

#[test]
fn columns_propagate_cell_error() {
    let mut cells = vec![
        XLOPER12::from_f64(1.0),
        XLOPER12::from_err(XLERR_NUM),
        XLOPER12::from_f64(3.0),
        XLOPER12::from_f64(4.0),
    ];
    let oper = make_multi(&mut cells, 2, 2);
    let err = xloper_to_columns(&oper).unwrap_err();
    assert_eq!(err, XLERR_NUM);
}

#[test]
fn columns_to_row_major_flattens() {
    let columns = vec![vec![2.0, 3.0], vec![5.0, 7.0]];
    let data = columns_to_row_major(&columns).unwrap();
    assert_eq!(data, vec![2.0, 5.0, 3.0, 7.0]);
}

#[test]
fn rows_to_row_major_flattens() {
    let rows = vec![vec![2.0, 5.0], vec![3.0, 7.0]];
    let data = rows_to_row_major(&rows).unwrap();
    assert_eq!(data, vec![2.0, 5.0, 3.0, 7.0]);
}

#[test]
#[should_panic(expected = "U type requires macro_equiv")]
fn build_type_string_panics_without_macro_equiv() {
    let _ = xll_rs::register::build_type_string('Q', &['U'], false, false, false);
}

#[test]
fn row_major_helpers_reject_ragged() {
    let columns = vec![vec![1.0, 2.0], vec![3.0]];
    assert_eq!(columns_to_row_major(&columns).unwrap_err(), XLERR_VALUE);

    let rows = vec![vec![1.0], vec![2.0, 3.0]];
    assert_eq!(rows_to_row_major(&rows).unwrap_err(), XLERR_VALUE);
}

#[test]
fn parse_optional_f64_variants() {
    let p = Box::into_raw(Box::new(XLOPER12::missing()));
    assert_eq!(parse_optional_f64(p, 2.0).unwrap(), 2.0);
    unsafe { free_xloper_recursive(p) };

    let p = Box::into_raw(Box::new(XLOPER12::nil()));
    assert_eq!(parse_optional_f64(p, 3.0).unwrap(), 3.0);
    unsafe { free_xloper_recursive(p) };

    let p = Box::into_raw(Box::new(XLOPER12::from_f64(1.5)));
    assert_eq!(parse_optional_f64(p, 0.0).unwrap(), 1.5);
    unsafe { free_xloper_recursive(p) };

    let p = Box::into_raw(Box::new(XLOPER12::from_int(4)));
    assert_eq!(parse_optional_f64(p, 0.0).unwrap(), 4.0);
    unsafe { free_xloper_recursive(p) };

    let p = Box::into_raw(Box::new(XLOPER12::from_err(XLERR_REF)));
    assert_eq!(parse_optional_f64(p, 0.0).unwrap_err(), XLERR_REF);
    unsafe { free_xloper_recursive(p) };
}

#[test]
fn parse_optional_bool_variants() {
    let p = Box::into_raw(Box::new(XLOPER12::missing()));
    assert_eq!(parse_optional_bool(p, true).unwrap(), true);
    unsafe { free_xloper_recursive(p) };

    let p = Box::into_raw(Box::new(XLOPER12::nil()));
    assert_eq!(parse_optional_bool(p, false).unwrap(), false);
    unsafe { free_xloper_recursive(p) };

    let p = Box::into_raw(Box::new(XLOPER12::from_bool(true)));
    assert_eq!(parse_optional_bool(p, false).unwrap(), true);
    unsafe { free_xloper_recursive(p) };

    let p = Box::into_raw(Box::new(XLOPER12::from_f64(0.0)));
    assert_eq!(parse_optional_bool(p, true).unwrap(), false);
    unsafe { free_xloper_recursive(p) };

    let p = Box::into_raw(Box::new(XLOPER12::from_int(2)));
    assert_eq!(parse_optional_bool(p, false).unwrap(), true);
    unsafe { free_xloper_recursive(p) };

    let p = Box::into_raw(Box::new(XLOPER12::from_err(XLERR_NUM)));
    assert_eq!(parse_optional_bool(p, false).unwrap_err(), XLERR_NUM);
    unsafe { free_xloper_recursive(p) };
}

#[test]
fn sheet_name_from_str_passes_through() {
    let oper = XLOPER12::from_str("Sheet1");
    let name = sheet_name_from_ref(&oper as *const _);
    assert_eq!(name.as_deref(), Some("Sheet1"));
}

#[test]
fn sheet_name_from_null_is_none() {
    let name = sheet_name_from_ref(std::ptr::null());
    assert!(name.is_none());
}

#[test]
fn coerce_to_owned_rejects_non_ref() {
    init_test_entrypoint();
    let oper = XLOPER12::from_f64(1.0);
    let err = coerce_to_owned(&oper).unwrap_err();
    assert_eq!(err, XLERR_VALUE);
}

#[test]
fn sheet_name_from_sref_uses_excel() {
    init_test_entrypoint();
    let sref = XLOPER12 {
        val: XLOPER12Val {
            sref: ManuallyDrop::new(XLOPER12SRef {
                count: 1,
                ref_: XLREF12 {
                    rw_first: 0,
                    rw_last: 0,
                    col_first: 0,
                    col_last: 0,
                },
            }),
        },
        xltype: XLTYPE_SREF,
    };
    let name = sheet_name_from_ref(&sref as *const _);
    assert_eq!(name.as_deref(), Some("TestSheet"));
}

#[test]
fn coerce_to_owned_from_sref() {
    init_test_entrypoint();
    let sref = XLOPER12 {
        val: XLOPER12Val {
            sref: ManuallyDrop::new(XLOPER12SRef {
                count: 1,
                ref_: XLREF12 {
                    rw_first: 0,
                    rw_last: 0,
                    col_first: 0,
                    col_last: 0,
                },
            }),
        },
        xltype: XLTYPE_SREF,
    };
    let p = coerce_to_owned(&sref).expect("coerce");
    unsafe {
        let oper = &*p;
        assert_eq!(oper.base_type(), XLTYPE_MULTI);
        let arr = &*std::ptr::addr_of!(oper.val.array);
        assert_eq!(arr.rows, 2);
        assert_eq!(arr.columns, 2);
        let cell0 = &*arr.lparray.add(0);
        let cell1 = &*arr.lparray.add(1);
        let cell2 = &*arr.lparray.add(2);
        let cell3 = &*arr.lparray.add(3);
        assert_eq!(cell0.as_f64(), Some(1.0));
        assert_eq!(cell1.as_f64(), Some(2.0));
        assert_eq!(cell2.as_f64(), Some(3.0));
        assert_eq!(cell3.as_f64(), Some(4.0));
    }
    xll_rs::memory::xlAutoFree12(p);
}
