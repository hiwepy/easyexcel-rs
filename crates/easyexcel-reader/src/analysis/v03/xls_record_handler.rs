//! Mirrors Java XLS record processing dispatch.

pub trait XlsRecordHandler {
    fn support(&self) -> bool { true }
    fn process_record(&mut self, record_sid: u16, data: &[u8]) {
        let _ = (record_sid, data);
    }
}
