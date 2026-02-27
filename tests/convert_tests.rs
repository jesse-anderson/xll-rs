use std::mem::ManuallyDrop;
use xll_rs::convert::{
    build_kv_table, build_multi, clone_excel_oper, return_xl_error, xloper_to_columns,
    xloper_to_f64_vec,
};
use xll_rs::memory::xlAutoFree12;
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
fn f64_vec_single_cell() {
    let oper = XLOPER12::from_f64(2.0);
    let v = xloper_to_f64_vec(&oper).expect("vec");
    assert_eq!(v, vec![2.0]);
}

#[test]
fn f64_vec_multi_cell() {
    let mut cells = vec![
        XLOPER12::from_f64(1.0),
        XLOPER12::from_f64(2.0),
        XLOPER12::from_f64(3.0),
        XLOPER12::from_f64(4.0),
    ];
    let oper = make_multi(&mut cells, 2, 2);
    let v = xloper_to_f64_vec(&oper).expect("vec");
    assert_eq!(v, vec![1.0, 2.0, 3.0, 4.0]);
}

#[test]
fn columns_from_multi() {
    // 2 rows x 2 cols: [[1,2],[3,4]] row-major
    let mut cells = vec![
        XLOPER12::from_f64(1.0),
        XLOPER12::from_f64(2.0),
        XLOPER12::from_f64(3.0),
        XLOPER12::from_f64(4.0),
    ];
    let oper = make_multi(&mut cells, 2, 2);
    let (cols, n_rows, n_cols) = xloper_to_columns(&oper).expect("cols");
    assert_eq!(n_rows, 2);
    assert_eq!(n_cols, 2);
    assert_eq!(cols[0], vec![1.0, 3.0]); // first column
    assert_eq!(cols[1], vec![2.0, 4.0]); // second column
}

#[test]
fn error_return_is_dll_free() {
    let p = return_xl_error(XLERR_NUM);
    assert!(!p.is_null());
    unsafe {
        assert_eq!((*p).base_type(), XLTYPE_ERR);
        assert_ne!((*p).xltype & XLBIT_DLL_FREE, 0);
    }
    unsafe { xlAutoFree12(p) };
}

#[test]
fn build_multi_and_free() {
    let cells = vec![
        XLOPER12::from_str("A"),
        XLOPER12::from_f64(1.0),
        XLOPER12::from_str("B"),
        XLOPER12::from_f64(2.0),
    ];
    let p = build_multi(cells, 2, 2);
    assert!(!p.is_null());
    unsafe { xlAutoFree12(p) };
}

#[test]
fn build_kv_table_two_columns() {
    let p = build_kv_table(&[("Mean", 1.5), ("Std", 2.5)]);
    assert!(!p.is_null());
    unsafe {
        let oper = &*p;
        assert_eq!(oper.base_type(), XLTYPE_MULTI);
        let arr = &*std::ptr::addr_of!(oper.val.array);
        assert_eq!(arr.rows, 2);
        assert_eq!(arr.columns, 2);

        let cell0 = &*arr.lparray.add(0);
        assert_eq!(cell0.as_string().as_deref(), Some("Mean"));
        let cell1 = &*arr.lparray.add(1);
        assert_eq!(cell1.as_f64(), Some(1.5));
        let cell2 = &*arr.lparray.add(2);
        assert_eq!(cell2.as_string().as_deref(), Some("Std"));
        let cell3 = &*arr.lparray.add(3);
        assert_eq!(cell3.as_f64(), Some(2.5));
    }
    unsafe { xlAutoFree12(p) };
}

#[test]
fn clone_excel_oper_scalars() {
    let num = XLOPER12::from_f64(3.5);
    let p = clone_excel_oper(&num).expect("clone num");
    unsafe {
        assert_eq!((*p).base_type(), XLTYPE_NUM);
        assert_eq!((*p).as_f64(), Some(3.5));
        assert_ne!((*p).xltype & XLBIT_DLL_FREE, 0);
    }
    unsafe { xlAutoFree12(p) };

    let s = XLOPER12::from_str("hi");
    let p = clone_excel_oper(&s).expect("clone str");
    unsafe {
        assert_eq!((*p).as_string().as_deref(), Some("hi"));
        assert_ne!((*p).xltype & XLBIT_DLL_FREE, 0);
    }
    unsafe { xlAutoFree12(p) };

    let b = XLOPER12::from_bool(true);
    let p = clone_excel_oper(&b).expect("clone bool");
    unsafe {
        assert_eq!((*p).base_type(), XLTYPE_BOOL);
        assert_eq!((*p).as_bool(), Some(true));
    }
    unsafe { xlAutoFree12(p) };
}

#[test]
fn clone_excel_oper_multi() {
    let cells = vec![
        XLOPER12::from_str("A"),
        XLOPER12::from_f64(1.0),
        XLOPER12::from_bool(true),
        XLOPER12::nil(),
    ];
    let p = build_multi(cells, 2, 2);
    let cloned = clone_excel_oper(unsafe { &*p }).expect("clone multi");
    unsafe {
        let oper = &*cloned;
        assert_eq!(oper.base_type(), XLTYPE_MULTI);
        let arr = &*std::ptr::addr_of!(oper.val.array);
        assert_eq!(arr.rows, 2);
        assert_eq!(arr.columns, 2);

        let cell0 = &*arr.lparray.add(0);
        assert_eq!(cell0.as_string().as_deref(), Some("A"));
        let cell1 = &*arr.lparray.add(1);
        assert_eq!(cell1.as_f64(), Some(1.0));
        let cell2 = &*arr.lparray.add(2);
        assert_eq!(cell2.as_bool(), Some(true));
        let cell3 = &*arr.lparray.add(3);
        assert_eq!(cell3.base_type(), XLTYPE_NIL);
    }
    unsafe { xlAutoFree12(p) };
    unsafe { xlAutoFree12(cloned) };
}

#[test]
fn clone_excel_oper_rejects_sref() {
    let oper = XLOPER12 {
        val: XLOPER12Val {
            sref: std::mem::ManuallyDrop::new(XLOPER12SRef {
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
    let err = clone_excel_oper(&oper).unwrap_err();
    assert_eq!(err, XLERR_VALUE);
}
