//! Mirrors Java `com.alibaba.excel.analysis.v07.handlers.AbstractCellValueTagHandler`.

use super::super::handlers::abstract_xlsx_tag_handler::AbstractXlsxTagHandler;

/// Mirrors Java `AbstractCellValueTagHandler extends AbstractXlsxTagHandler`.
///
/// Java's abstract class adds the `cellDataType` dispatch to concrete
/// value handlers. Rust inlines this logic into the SAX event match.
#[allow(dead_code)]
pub struct AbstractCellValueTagHandler;

impl super::xlsx_tag_handler::XlsxTagHandler for AbstractCellValueTagHandler {}
