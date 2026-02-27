//! Registry for xllgen-exported functions.

use crate::types::XLLError;

/// Matrix layout for Vec<Vec<T>> conversions.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum XllLayout {
    Row,
    Col,
}

/// Metadata for one exported XLL function.
#[derive(Clone, Copy, Debug)]
pub struct XllExport {
    pub rust_name: &'static str,
    pub base_name: &'static str,
    pub name: &'static str,
    pub auto_name: bool,
    pub aliases: &'static [&'static str],
    pub type_str: &'static str,
    pub arg_names: &'static str,
    pub category: &'static str,
    pub help: &'static str,
    pub arg_help: &'static [&'static str],
    pub threadsafe: bool,
    pub volatile: bool,
    pub layout: XllLayout,
    pub errors: &'static [XLLError],
}

inventory::collect!(XllExport);
