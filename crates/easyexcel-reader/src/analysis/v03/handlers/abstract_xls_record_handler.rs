//! Mirrors Java `com.alibaba.excel.analysis.v03.handlers.AbstractXlsRecordHandler`.

use super::super::xls_record_handler::XlsRecordHandler;

pub trait AbstractXlsRecordHandler: XlsRecordHandler {
    fn order(&self) -> i32 { 0 }
    fn process_record(&mut self, record_sid: u16, data: &[u8]) {
        let _ = (record_sid, data);
    }
}

pub type AbstractXlsRecordHandlerBox = Box<dyn AbstractXlsRecordHandler>;
