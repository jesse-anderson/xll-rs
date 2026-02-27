# scientific_xll

Example XLL add‑in built on `xll-rs` to pressure‑test the runtime with
scientific workflows and varied XLOPER12 usage.

## Build

```powershell
cd xll-rs\examples\scientific_xll
cargo build --release
```

Output:
```
target\release\scientific_xll.dll
```
Rename to `.xll` and load in Excel.

## Functions

- `SCI.VERSION()` — library version string
- `SCI.HELLO()` — greeting string
- `SCI.ADD(a, b)` — numeric add
- `SCI.NOT(value)` — boolean NOT
- `SCI.ECHO(text)` — string echo
- `SCI.MEAN(y_range)` — mean of a numeric range
- `SCI.DESCRIBE(y_range)` — 2‑column summary stats (mean/std/min/max)
- `SCI.CORR(y_range, x_range)` — correlation
- `SCI.COLSUM(x_range)` — column sums of matrix
- `SCI.SCALE(y_range, [factor], [center])` — optional args + scaling

- `SCI.NIL()` - return xltypeNil (Excel shows 0)
- `SCI.BLANK()` - return an empty string (blank cell)
- `SCI.ERROR(code)` - return an Excel error by code
- `SCI.TOINT(value)` - truncate a numeric value to int
- `SCI.THRESH(value, [cutoff], [strict])` - optional args + threshold comparison
- `SCI.TYPES()` - mixed-type table (string/number/bool/error/nil)
- `SCI.DIMS(x_range)` - rows/cols of numeric range
- `SCI.REFINFO(ref)` - reference metadata (SREF/REF)
- `SCI.REFVALUES(ref)` - coerce reference to values
- `SCI.FLOW()` - return xltypeFlow (testing)
- `SCI.BIGDATA()` - return xltypeBigData (Excel instance handle)

## Purpose

This is intentionally broad to exercise:
- scalar returns
- string returns
- boolean returns
- 1D and 2D arrays
- optional arguments
- error propagation
- memory cleanup via `xlAutoFree12`

## Excel Parity

`SCI_Workflow_Example.xlsx` includes end-to-end parity checks for every
function above (formulas + expected outputs).
