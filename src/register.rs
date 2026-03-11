//! UDF registration helper for `xlAutoOpen`.
//!
//! Wraps `xlfRegister` so each UDF can be registered with a single call.

use super::entrypoint::{excel12, excel_free, XLRET_SUCCESS};
use super::types::*;

/// Registrator — caches the DLL path and provides a simple `add()` method.
pub struct Reg {
    dll_name: XLOPER12,
}

/// Build a comma-separated argument name list for `xlfRegister`.
pub fn build_arg_names(names: &[&str]) -> String {
    names.join(", ")
}

/// Build a type string for `xlfRegister`.
///
/// `return_type` and `arg_types` are the type characters (e.g. 'Q', 'B', 'J').
/// Modifiers are appended after arguments:
/// - `#` macro sheet equivalent (required for `U`)
/// - `$` thread-safe
/// - `!` volatile
///
/// Excel does not allow `#` and `$` together; if `macro_equiv` is true, `$` is omitted.
pub fn build_type_string(
    return_type: char,
    arg_types: &[char],
    threadsafe: bool,
    volatile: bool,
    macro_equiv: bool,
) -> String {
    assert!(
        !(arg_types.iter().any(|c| *c == 'U') && !macro_equiv),
        "U type requires macro_equiv (#) in the type string"
    );
    let mut out = String::with_capacity(1 + arg_types.len() + 2);
    out.push(return_type);
    for t in arg_types {
        out.push(*t);
    }
    if macro_equiv {
        out.push('#');
    } else if threadsafe {
        out.push('$');
    }
    if volatile {
        out.push('!');
    }
    out
}

impl Reg {
    /// Create a new registrator.  Fetches the DLL path from Excel via
    /// `xlGetName`.
    pub fn new() -> Self {
        let (_, dll_name) = excel12(XL_GET_NAME, &mut []);
        Reg { dll_name }
    }

    /// Register one UDF with Excel.
    ///
    /// # Arguments
    ///
    /// * `fn_name` — Exported C symbol name (e.g. `"xl_version"`)
    /// * `type_str` — Type string: return type char + arg type chars + modifiers.
    ///   Common codes: `Q` = XLOPER12 by ref, `U` = XLOPER12 by ref (no coercion),
    ///   `B` = f64, `J` = i32. Modifiers: `$` = thread-safe, `!` = volatile,
    ///   `#` = macro sheet equivalent (required for `U`).
    /// * `excel_name` — Name shown in Excel (e.g. `"LINREG.VERSION"`)
    /// * `arg_names` — Comma-separated argument names for the Function Wizard
    /// * `category` — Category in the Function Wizard (e.g. `"LinReg"`)
    /// * `description` — Short help text shown in the Function Wizard
    /// * `arg_help` — Per-argument help strings (xlfRegister args 11+).
    ///   Pass `&[]` for no per-argument help.
    pub fn add(
        &self,
        fn_name: &str,
        type_str: &str,
        excel_name: &str,
        arg_names: &str,
        category: &str,
        description: &str,
        arg_help: &[&str],
    ) -> Result<(), i32> {
        let mut dll = self.dll_name_ptr();
        let mut fn_n = XLOPER12::from_str(fn_name);
        let mut types = XLOPER12::from_str(type_str);
        let mut xl_n = XLOPER12::from_str(excel_name);
        let mut args = XLOPER12::from_str(arg_names);
        let mut macro_type = XLOPER12::from_int(1); // 1 = worksheet function
        let mut cat = XLOPER12::from_str(category);
        let mut shortcut = XLOPER12::missing();
        let mut help_topic = XLOPER12::missing();
        let mut desc = XLOPER12::from_str(description);

        let mut arg_help_opers: Vec<XLOPER12> =
            arg_help.iter().map(|s| XLOPER12::from_str(s)).collect();

        let mut opers: Vec<*mut XLOPER12> = Vec::with_capacity(10 + arg_help_opers.len());
        opers.extend_from_slice(&mut [
            &mut dll as *mut _,
            &mut fn_n,
            &mut types,
            &mut xl_n,
            &mut args,
            &mut macro_type,
            &mut cat,
            &mut shortcut,
            &mut help_topic,
            &mut desc,
        ]);
        for h in arg_help_opers.iter_mut() {
            opers.push(h);
        }

        let (ret, res) = excel12(XLF_REGISTER, &mut opers);

        // If Excel returned an XLOPER12 we need to free, do so
        let mut res = res;
        if (res.xltype & XLBIT_XL_FREE) != 0 {
            excel_free(&mut res);
        }

        // Free the temp XLOPER12 strings we allocated
        free_xloper(&mut fn_n);
        free_xloper(&mut types);
        free_xloper(&mut xl_n);
        free_xloper(&mut args);
        free_xloper(&mut cat);
        free_xloper(&mut desc);
        for h in arg_help_opers.iter_mut() {
            free_xloper(h);
        }

        if ret == XLRET_SUCCESS {
            Ok(())
        } else {
            Err(ret)
        }
    }

    fn dll_name_ptr(&self) -> XLOPER12 {
        // Return a shallow copy without ownership flags — we don't want
        // the registration call to free our cached dll_name.
        XLOPER12 {
            val: XLOPER12Val {
                str_: unsafe { self.dll_name.val.str_ },
            },
            xltype: XLTYPE_STR, // no DLLFree bit
        }
    }
}

impl Drop for Reg {
    fn drop(&mut self) {
        // dll_name was returned by Excel (xlGetName) — free via xlFree,
        // not free_xloper, because Excel owns the underlying memory.
        excel_free(&mut self.dll_name);
    }
}

/// Free a DLL-owned XLOPER12's string/array memory.
///
/// This helper only frees strings allocated by `XLOPER12::from_str` and is
/// intended for cleaning up registration temporaries. For general XLOPER12
/// cleanup, use `memory::xlAutoFree12`.
fn free_xloper(oper: &mut XLOPER12) {
    if (oper.xltype & XLBIT_DLL_FREE) == 0 {
        return;
    }
    let base = oper.xltype & 0x0FFF;
    match base {
        XLTYPE_STR => unsafe {
            let ptr = oper.val.str_;
            if !ptr.is_null() {
                let len = *ptr as usize + 1;
                let _ = Vec::from_raw_parts(ptr, len, len);
            }
        },
        _ => {}
    }
    oper.xltype = XLTYPE_NIL;
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::entrypoint::init_test_entrypoint;
    use crate::types::*;

    #[test]
    fn build_arg_names_joins() {
        assert_eq!(build_arg_names(&[]), "");
        assert_eq!(build_arg_names(&["a"]), "a");
        assert_eq!(build_arg_names(&["a", "b", "c"]), "a, b, c");
    }

    #[test]
    fn build_type_string_formats_modifiers() {
        let s = build_type_string('Q', &['U'], true, false, true);
        assert_eq!(s, "QU#");

        let s = build_type_string('Q', &['Q', 'Q'], true, true, false);
        assert_eq!(s, "QQQ$!");

        let s = build_type_string('Q', &[], false, false, false);
        assert_eq!(s, "Q");
    }

    #[test]
    fn free_xloper_releases_string() {
        let mut oper = XLOPER12::from_str("free");
        free_xloper(&mut oper);
        assert_eq!(oper.xltype, XLTYPE_NIL);
    }

    #[test]
    fn reg_add_returns_result_code() {
        init_test_entrypoint();
        let reg = Reg {
            dll_name: XLOPER12::nil(),
        };
        let ret = reg.add(
            "xl_test",
            "Q$",
            "TEST",
            "",
            "Test",
            "Test",
            &[],
        );
        assert!(ret.is_ok());
    }
}
