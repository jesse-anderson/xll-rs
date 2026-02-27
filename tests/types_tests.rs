use xll_rs::types::*;

#[test]
fn string_roundtrip() {
    let oper = XLOPER12::from_str("hello");
    assert_eq!(oper.base_type(), XLTYPE_STR);
    assert_ne!(oper.xltype & XLBIT_DLL_FREE, 0);
    let s = oper.as_string().expect("string decode");
    assert_eq!(s, "hello");
}

#[test]
fn numeric_and_bool() {
    let n = XLOPER12::from_f64(3.5);
    assert_eq!(n.base_type(), XLTYPE_NUM);
    assert_eq!(n.as_f64(), Some(3.5));

    let i = XLOPER12::from_int(7);
    assert_eq!(i.base_type(), XLTYPE_INT);
    assert_eq!(i.as_f64(), Some(7.0));

    let b = XLOPER12::from_bool(true);
    assert_eq!(b.base_type(), XLTYPE_BOOL);
    assert_eq!(b.as_bool(), Some(true));
}
