//! Mirrors Java `com.alibaba.excel.analysis.v03.handlers.AbstractXlsRecordHandler`.

use super::super::xls_record_handler::XlsRecordHandler;

/// Rust marker counterpart of Java's abstract base class.
///
/// It cannot be instantiated and therefore cannot become an accidental no-op
/// record handler. Concrete handlers must implement [`XlsRecordHandler`].
pub trait AbstractXlsRecordHandler: XlsRecordHandler {}

macro_rules! impl_abstract_handler {
    ($($handler:path),+ $(,)?) => {
        $(impl AbstractXlsRecordHandler for $handler {})+
    };
}

impl_abstract_handler!(
    super::blank_record_handler::BlankRecordHandler,
    super::bof_record_handler::BofRecordHandler,
    super::bool_err_record_handler::BoolErrRecordHandler,
    super::bound_sheet_record_handler::BoundSheetRecordHandler,
    super::dummy_record_handler::DummyRecordHandler,
    super::eof_record_handler::EofRecordHandler,
    super::formula_record_handler::FormulaRecordHandler,
    super::hyperlink_record_handler::HyperlinkRecordHandler,
    super::index_record_handler::IndexRecordHandler,
    super::label_record_handler::LabelRecordHandler,
    super::label_sst_record_handler::LabelSstRecordHandler,
    super::merge_cells_record_handler::MergeCellsRecordHandler,
    super::note_record_handler::NoteRecordHandler,
    super::number_record_handler::NumberRecordHandler,
    super::obj_record_handler::ObjRecordHandler,
    super::rk_record_handler::RkRecordHandler,
    super::sst_record_handler::SstRecordHandler,
    super::string_record_handler::StringRecordHandler,
    super::text_object_record_handler::TextObjectRecordHandler,
);
