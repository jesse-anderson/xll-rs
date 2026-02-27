use std::path::PathBuf;

/// Emit a native `.xll` output on Windows/MSVC builds.
///
/// Call this from your XLL crate's `build.rs`:
/// `xll_rs::build::emit_xll();`
pub fn emit_xll() {
    let target_os = std::env::var("CARGO_CFG_TARGET_OS").unwrap_or_default();
    if target_os != "windows" {
        panic!("xll-rs: Excel XLLs require Windows. Use --target x86_64-pc-windows-msvc.");
    }

    let target_env = std::env::var("CARGO_CFG_TARGET_ENV").unwrap_or_default();
    if target_env != "msvc" {
        panic!("xll-rs: MSVC toolchain required. Use --target x86_64-pc-windows-msvc.");
    }

    let manifest_dir =
        PathBuf::from(std::env::var("CARGO_MANIFEST_DIR").unwrap_or_else(|_| ".".to_string()));
    let profile = std::env::var("PROFILE").unwrap_or_else(|_| "debug".to_string());
    let target_dir = match std::env::var("CARGO_TARGET_DIR") {
        Ok(val) => {
            let p = PathBuf::from(val);
            if p.is_relative() {
                manifest_dir.join(p)
            } else {
                p
            }
        }
        Err(_) => manifest_dir.join("target"),
    };
    let name = std::env::var("CARGO_PKG_NAME").unwrap_or_else(|_| "xll".to_string());
    let out = target_dir.join(&profile).join(format!("{name}.xll"));

    // Force link.exe to emit .xll directly. The /OUT value wins over rustc's -o.
    println!("cargo:rustc-cdylib-link-arg=/OUT:{}", out.display());
}
