//! Mirrors Java `com.alibaba.excel.analysis.v03.IgnorableXlsRecordHandler`.

use super::xls_record_handler::XlsRecordHandler;

/// Mirrors Java `IgnorableXlsRecordHandler extends XlsRecordHandler`.
///
/// Java marks handlers whose records belong to a worksheet and may be skipped
/// while the current worksheet is not selected.
pub trait IgnorableXlsRecordHandler: XlsRecordHandler {}

macro_rules! impl_ignorable {
    ($($handler:path),+ $(,)?) => {
        $(impl IgnorableXlsRecordHandler for $handler {})+
    };
}

impl_ignorable!(
    super::handlers::blank_record_handler::BlankRecordHandler,
    super::handlers::bool_err_record_handler::BoolErrRecordHandler,
    super::handlers::bound_sheet_record_handler::BoundSheetRecordHandler,
    super::handlers::dummy_record_handler::DummyRecordHandler,
    super::handlers::formula_record_handler::FormulaRecordHandler,
    super::handlers::hyperlink_record_handler::HyperlinkRecordHandler,
    super::handlers::index_record_handler::IndexRecordHandler,
    super::handlers::label_record_handler::LabelRecordHandler,
    super::handlers::label_sst_record_handler::LabelSstRecordHandler,
    super::handlers::merge_cells_record_handler::MergeCellsRecordHandler,
    super::handlers::note_record_handler::NoteRecordHandler,
    super::handlers::number_record_handler::NumberRecordHandler,
    super::handlers::obj_record_handler::ObjRecordHandler,
    super::handlers::rk_record_handler::RkRecordHandler,
    super::handlers::sst_record_handler::SstRecordHandler,
    super::handlers::string_record_handler::StringRecordHandler,
    super::handlers::text_object_record_handler::TextObjectRecordHandler,
);
