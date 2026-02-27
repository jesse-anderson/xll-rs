//! xll-rs — Excel XLL runtime for Rust.
//!
//! Windows-only crate. Excel XLLs are not supported on other platforms.

#[cfg(not(windows))]
compile_error!("xll-rs is Windows-only (Excel XLLs require Windows/MSVC).");

pub mod types;
pub mod entrypoint;
pub mod register;
pub mod convert;
pub mod build;
pub mod memory;
pub mod returning;
pub mod registry;

// Re-export inventory so xllgen users don't need to depend on it directly.
pub use inventory;
