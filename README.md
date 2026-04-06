# xll-rs

Rust runtime for Excel XLL add-ins. Provides:
- XLOPER12 types and constants
- Excel12v trampoline (MdCallBack12)
- UDF registration helpers
- Conversion helpers (ranges to `Vec<f64>`, column-major extraction, design matrices)
- Memory ownership helpers (`xlAutoFree12`)
- Return value wrapper (`XlReturn`) that guarantees `xlbitDLLFree`

**Windows-only:** Excel XLLs run on Windows. Non-Windows targets emit a compile-time error.

## Modules

| Module | Description |
|--------|-------------|
| `types` | XLOPER12 struct, union variants, xltype constants, error codes, C API function numbers |
| `entrypoint` | MdCallBack12 trampoline (`excel12v`, `excel12`), `excel_free` |
| `register` | `Reg` helper for `xlfRegister` — caches DLL path, builds argument arrays |
| `convert` | `xloper_to_f64_vec`, `xloper_to_columns`, `build_multi`, `return_xl_error`, `columns_to_row_major`, `rows_to_row_major`, `parse_optional_f64`, `parse_optional_bool` |
| `memory` | `xlAutoFree12` — recursive freeing of DLL-owned XLOPER12s (strings, arrays) |
| `returning` | `XlReturn` — owned wrapper that sets `xlbitDLLFree` and heap-allocates via `Box::into_raw` |
| `build` | Build helper for emitting `.xll` output from `build.rs` |

## Native .xll Output (No Renaming)

Add a `build.rs` in your XLL crate:

```rust
fn main() {
    xll_rs::build::emit_xll();
}
```

This emits `target/<profile>/<crate-name>.xll` on every build.
Requires Windows + MSVC toolchain (`x86_64-pc-windows-msvc`).

## XLOPER12 Types

The XLOPER12 is a tagged union. The `xltype` field determines which union variant is active.

### Value Types

| Constant | Value | Union field | Rust type | Description |
|----------|-------|-------------|-----------|-------------|
| `XLTYPE_NUM` | `0x0001` | `val.num` | `f64` | 64-bit floating point |
| `XLTYPE_STR` | `0x0002` | `val.str_` | `*mut u16` | Length-counted UTF-16 string (Pascal-style: `[len, char0, char1, ...]`) |
| `XLTYPE_BOOL` | `0x0004` | `val.xbool` | `i32` | Boolean (0 = false, 1 = true) |
| `XLTYPE_REF` | `0x0008` | `val.mref` | `XLOPER12MRef` | Multi-area or cross-sheet cell reference (sheet ID + reference table) |
| `XLTYPE_ERR` | `0x0010` | `val.err` | `i32` | Error value (see error codes below) |
| `XLTYPE_FLOW` | `0x0020` | `val.flow` | `XLOPER12Flow` | Macro flow control (halt, goto, restart, pause, resume) |
| `XLTYPE_MULTI` | `0x0040` | `val.array` | `XLOPER12Array` | 2D array of XLOPER12s (row-major, rows x columns) |
| `XLTYPE_MISSING` | `0x0080` | `val.w` | `i32` | Omitted argument (caller did not provide this parameter) |
| `XLTYPE_NIL` | `0x0100` | `val.w` | `i32` | Empty cell / no value |
| `XLTYPE_SREF` | `0x0400` | `val.sref` | `XLOPER12SRef` | Same-sheet single-area cell reference |
| `XLTYPE_INT` | `0x0800` | `val.w` | `i32` | 32-bit signed integer |
| `XLTYPE_BIGDATA` | `0x0802` | `val.bigdata` | `XLOPER12BigData` | Binary data or Excel instance handle (xltypeStr \| xltypeInt) |

### Memory Flag Bits

| Constant | Value | Description |
|----------|-------|-------------|
| `XLBIT_XL_FREE` | `0x1000` | Excel owns this memory. Call `excel_free()` (which invokes `xlFree`) to release it. Set by Excel on return values from C API calls like `xlGetName`, `xlSheetNm`, `xlCoerce`. |
| `XLBIT_DLL_FREE` | `0x4000` | The DLL owns this memory. Excel will call `xlAutoFree12` after copying the result. Set by `XlReturn::from_oper()`, `XLOPER12::from_str()`, and `build_multi()`. |

### Error Codes

| Constant | Value | Excel display |
|----------|-------|---------------|
| `XLERR_NULL` | `0` | `#NULL!` |
| `XLERR_DIV0` | `7` | `#DIV/0!` |
| `XLERR_VALUE` | `15` | `#VALUE!` |
| `XLERR_REF` | `23` | `#REF!` |
| `XLERR_NAME` | `29` | `#NAME?` |
| `XLERR_NUM` | `36` | `#NUM!` |
| `XLERR_NA` | `42` | `#N/A` |
| `XLERR_GETTING_DATA` | `43` | `#GETTING_DATA` |

### Flow Control Codes

Used with `XLTYPE_FLOW`. These are XLM macro control operations.

| Constant | Value |
|----------|-------|
| `XLFLOW_HALT` | `1` |
| `XLFLOW_GOTO` | `2` |
| `XLFLOW_RESTART` | `8` |
| `XLFLOW_PAUSE` | `16` |
| `XLFLOW_RESUME` | `64` |

### Excel C API Function Numbers

| Constant | Value | Description |
|----------|-------|-------------|
| `XL_FREE` | `0x4000` | Free Excel-owned XLOPER12 memory |
| `XL_GET_NAME` | `0x4009` | Get the DLL file path |
| `XLF_REGISTER` | `149` | Register a UDF with Excel |
| `XL_COERCE` | `0x4002` | Coerce an XLOPER12 to a different type (e.g. reference to values) |
| `XL_SHEET_ID` | `0x4004` | Get sheet ID from name |
| `XL_SHEET_NM` | `0x4005` | Get sheet name from reference |
| `XL_GET_INST` | `0x4007` | Get Excel HINSTANCE |
| `XL_GET_INST_PTR` | `0x4013` | Get Excel instance handle as BigData (Excel 2010+) |

## Registration Type Strings

The `type_text` parameter to `xlfRegister` encodes the return type, argument types, and modifiers.

### Data Type Characters

| Character | Type | Description |
|-----------|------|-------------|
| `Q` | `XLOPER12*` | Passed by reference. Excel coerces cell references to values automatically. You receive `xltypeNum`, `xltypeStr`, `xltypeBool`, `xltypeErr`, `xltypeMulti`, `xltypeMissing`, or `xltypeNil` -- never `xltypeRef` or `xltypeSRef`. |
| `U` | `XLOPER12*` | Passed by reference without coercion. You receive the raw reference (`xltypeSRef` or `xltypeRef`). Requires the `#` modifier. |
| `B` | `f64` | Double-precision floating point, passed by value. |
| `J` | `i32` | 32-bit signed integer, passed by value. |
| `A` | `i16` (bool) | Boolean, passed by value (0 = false, 1 = true). |

### Modifier Characters

Appended after the last argument type character.

| Modifier | Description |
|----------|-------------|
| `$` | Thread-safe. Excel may call the function from multiple threads simultaneously. Cannot be combined with `#`. |
| `#` | Macro sheet equivalent. Grants the function permission to call XLM information functions (`xlSheetNm`, `xlCoerce`, `xlfGetCell`, etc.) and to receive raw references via `U` type. Cannot be combined with `$`. Functions with `#` are volatile by default when using `U` type arguments. |
| `!` | Volatile. The function recalculates every time the worksheet recalculates, regardless of whether its inputs changed. |

### Type String Format

The first character is the return type. Remaining characters (before modifiers) are argument types.

| Example | Meaning |
|---------|---------|
| `Q$` | Returns `XLOPER12*`, no arguments, thread-safe |
| `QQ$` | Returns `XLOPER12*`, one `Q` argument, thread-safe |
| `QQQ$` | Returns `XLOPER12*`, two `Q` arguments, thread-safe |
| `QQQQ$` | Returns `XLOPER12*`, three `Q` arguments, thread-safe |
| `QU#` | Returns `XLOPER12*`, one `U` argument (raw reference), macro sheet equivalent |
| `Q` | Returns `XLOPER12*`, no arguments, no modifiers |

Helpers in `xll_rs::register`: `build_type_string(...)` and `build_arg_names(...)` provide safe defaults for building
`type_text` and comma-separated argument name lists.

## Constructors

### Creating XLOPER12 Values

```rust
XLOPER12::from_f64(3.14)       // xltypeNum
XLOPER12::from_str("hello")    // xltypeStr (allocates UTF-16 buffer, sets xlbitDLLFree)
XLOPER12::from_int(42)         // xltypeInt
XLOPER12::from_bool(true)      // xltypeBool
XLOPER12::from_err(XLERR_NA)   // xltypeErr
XLOPER12::nil()                // xltypeNil
XLOPER12::missing()            // xltypeMissing
```

### XlReturn Wrapper

`XlReturn` wraps an XLOPER12, guarantees `xlbitDLLFree` is set, and heap-allocates
via `into_raw()` for returning to Excel:

```rust
XlReturn::num(3.14).into_raw()     // -> *mut XLOPER12
XlReturn::str("hello").into_raw()
XlReturn::int(42).into_raw()
XlReturn::bool(true).into_raw()
XlReturn::err(XLERR_NA).into_raw()
XlReturn::nil().into_raw()
```

### Multi Arrays

```rust
let mut cells = vec![
    XLOPER12::from_str("Label"),
    XLOPER12::from_f64(42.0),
];
build_multi(cells, 1, 2)  // 1 row, 2 columns -> *mut XLOPER12
```

## Minimal Example

```rust
use xll_rs::register::Reg;
use xll_rs::returning::XlReturn;
use xll_rs::types::*;

#[no_mangle]
pub extern "system" fn xl_hello() -> *mut XLOPER12 {
    XlReturn::str("Hello from Rust").into_raw()
}

#[no_mangle]
pub extern "system" fn xlAutoOpen() -> i32 {
    let reg = Reg::new();
    reg.add(
        "xl_hello",
        "Q$",
        "HELLO.RUST",
        "",
        "User",
        "Returns a greeting string",
        &[],
    );
    1
}

// Excel calls this after it copies results with xlbitDLLFree set
pub use xll_rs::memory::xlAutoFree12;
```

## scientific_xll Example

The `examples/scientific_xll` crate demonstrates every XLOPER12 type and registration
pattern. Below is each UDF grouped by what it exercises.

### Scalar Returns

| UDF | Type string | Types exercised | Description |
|-----|-------------|----------------|-------------|
| `SCI.VERSION` | `Q$` | Returns **xltypeStr** | Library version string via `env!("CARGO_PKG_VERSION")` |
| `SCI.HELLO` | `Q$` | Returns **xltypeStr** | Static greeting string |
| `SCI.ADD` | `QQQ$` | Reads **xltypeNum** from two Q args, returns **xltypeNum** | Adds two numbers |
| `SCI.NOT` | `QQ$` | Reads **xltypeBool**, returns **xltypeBool** | Logical NOT |
| `SCI.ECHO` | `QQ$` | Reads **xltypeStr**, returns **xltypeStr** | Echoes input string back |
| `SCI.MEAN` | `QQ$` | Reads **xltypeMulti** (coerced from range), returns **xltypeNum** | Mean of a numeric range |
| `SCI.CORR` | `QQQ$` | Reads two **xltypeMulti** ranges, returns **xltypeNum** | Pearson correlation |
| `SCI.TOINT` | `QQ$` | Reads **xltypeNum**, returns **xltypeInt** | Truncates float to integer |
| `SCI.NIL` | `Q$` | Returns **xltypeNil** | Returns an empty/nil cell |
| `SCI.BLANK` | `Q$` | Returns **xltypeStr** (empty) | Returns zero-length string |
| `SCI.REGDIAG` | `Q$` | Returns **xltypeStr** | Registration diagnostics from `xlAutoOpen` |

### Optional Arguments (xltypeMissing)

| UDF | Type string | Types exercised | Description |
|-----|-------------|----------------|-------------|
| `SCI.SCALE` | `QQQQ$` | Reads **xltypeMissing** for optional args, **xltypeNum** for factor, **xltypeBool** for center flag | Scale a vector with optional centering |
| `SCI.THRESH` | `QQQQ$` | Reads **xltypeMissing** for optional cutoff/strict args | Threshold comparison with defaults |

### Multi (Array) Returns

| UDF | Type string | Types exercised | Description |
|-----|-------------|----------------|-------------|
| `SCI.DESCRIBE` | `QQ$` | Returns **xltypeMulti** containing **xltypeStr** labels + **xltypeNum** values | Summary stats table (mean, std, min, max) |
| `SCI.COLSUM` | `QQ$` | Returns **xltypeMulti** with mixed **xltypeStr** and **xltypeNum** | Column sums of a matrix |
| `SCI.DIMS` | `QQ$` | Returns **xltypeMulti** (Rows/Cols) | Dimensions of a numeric range |
| `SCI.TYPES` | `Q$` | Returns **xltypeMulti** containing every displayable type: **xltypeStr**, **xltypeNum**, **xltypeInt**, **xltypeBool**, all 7 **xltypeErr** codes, **xltypeNil**, **xltypeMissing** | Type showcase table |

### Error Returns

| UDF | Type string | Types exercised | Description |
|-----|-------------|----------------|-------------|
| `SCI.ERROR` | `QQ$` | Reads error code, returns **xltypeErr** | Returns a specific Excel error by numeric code |

All UDFs also return **xltypeErr** on invalid input (null pointers, wrong types, computation failures).

### Reference Types (U type, macro sheet equivalent)

| UDF | Type string | Types exercised | Description |
|-----|-------------|----------------|-------------|
| `SCI.REFINFO` | `QU#` | Reads **xltypeSRef** (same-sheet ref) or **xltypeRef** (multi-area/cross-sheet ref). Calls **xlSheetNm** to get sheet name. Returns **xltypeMulti**. | Displays reference metadata: type, sheet name, sheet ID, area coordinates |
| `SCI.REFVALUES` | `QU#` | Reads **xltypeSRef**/**xltypeRef**, calls **xlCoerce** to convert to **xltypeMulti**, then deep-clones into DLL-owned memory | Coerces a reference to its values |

### Exotic Types

| UDF | Type string | Types exercised | Description |
|-----|-------------|----------------|-------------|
| `SCI.FLOW` | `Q` | Constructs and returns **xltypeFlow** with `XLFLOW_RESTART` | Tests flow control type (XLM macro concept) |
| `SCI.BIGDATA` | `Q` | Calls **xlGetInstPtr**, reads **xltypeBigData** result, returns it | Tests BigData type (Excel 2010+ instance handle) |

### Excel Callbacks

| Export | Description |
|--------|-------------|
| `xlAutoOpen` | Called when the XLL loads. Creates a `Reg` (fetches DLL path via `xlGetName`), registers all UDFs via `xlfRegister`. |
| `xlAutoClose` | Called when the XLL unloads. Returns 1 (no cleanup needed). |
| `xlAutoFree12` | Called by Excel when it finishes with a returned XLOPER12 that has `xlbitDLLFree` set. Recursively frees strings and arrays. Re-exported from `xll_rs::memory`. |
| `xlAddInManagerInfo12` | Returns the add-in display name for the Add-In Manager dialog. |

Downstream crates can re-export `xll_rs::memory::xlAutoFree12` directly (as shown in the minimal example).

### Helper Functions

| Function | Types handled | Description |
|----------|--------------|-------------|
| `parse_optional_f64` | xltypeMissing, xltypeNil, xltypeNum, xltypeInt, xltypeErr | Extracts an f64 from an optional argument, returning a default if missing |
| `parse_optional_bool` | xltypeMissing, xltypeNil, xltypeBool, xltypeNum, xltypeInt, xltypeErr | Extracts a bool from an optional argument, returning a default if missing |
| `build_kv_table` | xltypeStr, xltypeNum | Builds a 2-column label/value multi array |
| `sheet_name_from_ref` | xltypeSRef, xltypeRef, xltypeStr | Calls xlSheetNm on a reference (or passes through strings) and frees Excel-owned result |
| `clone_excel_oper` | xltypeNum, xltypeInt, xltypeBool, xltypeErr, xltypeNil, xltypeStr, xltypeMulti | Deep-clones an Excel-owned XLOPER12 into DLL-owned memory |
| `coerce_to_owned` | xltypeSRef, xltypeRef -> xltypeMulti | Calls xlCoerce to dereference a cell reference, then deep-clones the result |

## Developer Workflow: Reloading an XLL in Excel

Excel locks the `.xll` file while it is loaded, and caches add-ins between sessions.
Simply overwriting the file and reopening Excel will often load the stale copy.

To force Excel to pick up a new build:

1. **File > Options > Add-ins > Manage: Excel Add-ins > Go...**
2. Uncheck the XLL entry and click OK.
3. Return to the same dialog, select the entry, and click **Remove**.
4. Close Excel completely (verify no `EXCEL.EXE` remains in Task Manager).
5. Copy/overwrite the `.xll` file now that the lock is released.
6. Reopen Excel and re-add the XLL via **Browse**.

If Excel is still loading the old version after this process, rename the file to break
the cache entirely (e.g. `scientific_xll_1.xll`, `scientific_xll_2.xll`). This forces
Excel to treat it as a new add-in with no cached state. Remove the old entry afterward.

