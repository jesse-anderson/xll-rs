# Changelog

## [0.1.1] - 2026-03-11

### Changed
- `Reg::add()` now returns `Result<(), i32>` instead of raw `i32`, preventing
  misinterpretation of the success return code

## [0.1.0] - 2026-02-26

### Added
- XLOPER12 types and constants (including flow and bigdata variants)
- Excel12v trampoline via `MdCallBack12`
- UDF registration helper (`Reg`)
- Conversion helpers (`xloper_to_f64_vec`, `xloper_to_columns`, `build_multi`)
- Memory ownership helpers and `xlAutoFree12`
