//! Mirrors Java `com.alibaba.excel.analysis.v03.*`.

pub(crate) mod biff_record_stream;
pub(crate) mod biff_string;
pub mod handlers;
pub mod ignorable_xls_record_handler;
pub mod xls_list_sheet_listener;
pub mod xls_record_dispatcher;
pub mod xls_record_handler;
pub mod xls_sax_analyser;

pub use xls_record_dispatcher::{XlsRecordDispatchState, XlsRecordDispatcher};
pub use xls_sax_analyser::XlsSaxAnalyser;
