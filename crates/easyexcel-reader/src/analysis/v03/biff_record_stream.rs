//! Low-level BIFF record stream used by the XLS event compatibility layer.

use std::fs::File;
use std::io::Read;
use std::path::Path;

use cfb::CompoundFile;
use easyexcel_core::{ExcelError, Result};

/// Reads the BIFF workbook stream from an OLE2 `.xls` compound document.
pub(crate) fn read_workbook_stream(path: &Path) -> Result<Vec<u8>> {
    let file = File::open(path)?;
    let mut compound = CompoundFile::open(file)
        .map_err(|error| ExcelError::Format(format!("invalid XLS compound document: {error}")))?;
    let mut stream = compound
        .open_stream("/Workbook")
        .or_else(|_| compound.open_stream("/Book"))
        .map_err(|error| {
            ExcelError::Format(format!("XLS Workbook/Book stream is missing: {error}"))
        })?;
    let mut workbook = Vec::new();
    stream.read_to_end(&mut workbook)?;
    Ok(workbook)
}

/// Walks every physical BIFF record in a workbook stream.
///
/// Unlike the former display-only parser, this reports truncated headers and
/// payloads instead of silently accepting a damaged stream.
pub(crate) fn walk_biff_records(
    workbook: &[u8],
    mut process: impl FnMut(u16, &[u8]) -> Result<()>,
) -> Result<()> {
    let mut offset = 0usize;
    while offset < workbook.len() {
        let remaining = &workbook[offset..];
        if remaining.iter().all(|byte| *byte == 0) {
            break;
        }
        if remaining.len() < 4 {
            return Err(ExcelError::Format(format!(
                "truncated BIFF record header at byte {offset}"
            )));
        }

        let sid = u16::from_le_bytes([remaining[0], remaining[1]]);
        let length = u16::from_le_bytes([remaining[2], remaining[3]]) as usize;
        let payload_start = offset + 4;
        let payload_end = payload_start.checked_add(length).ok_or_else(|| {
            ExcelError::Format(format!("BIFF record length overflow at byte {offset}"))
        })?;
        if payload_end > workbook.len() {
            return Err(ExcelError::Format(format!(
                "truncated BIFF record 0x{sid:04X} at byte {offset}: expected {length} payload bytes"
            )));
        }

        process(sid, &workbook[payload_start..payload_end])?;
        offset = payload_end;
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn walks_records_in_physical_order() -> Result<()> {
        let bytes = [0x03, 0x02, 0x02, 0x00, 0xAA, 0xBB, 0x0A, 0x00, 0x00, 0x00];
        let mut records = Vec::new();
        walk_biff_records(&bytes, |sid, payload| {
            records.push((sid, payload.to_vec()));
            Ok(())
        })?;
        assert_eq!(records, vec![(0x0203, vec![0xAA, 0xBB]), (0x000A, vec![])]);
        Ok(())
    }

    #[test]
    fn rejects_truncated_payload() {
        let error = walk_biff_records(&[0x03, 0x02, 0x04, 0x00, 0xAA], |_, _| Ok(()))
            .expect_err("payload is truncated");
        assert!(error.to_string().contains("truncated BIFF record 0x0203"));
    }

    #[test]
    fn ignores_zero_padding() -> Result<()> {
        let mut seen = 0;
        walk_biff_records(&[0; 16], |_, _| {
            seen += 1;
            Ok(())
        })?;
        assert_eq!(seen, 0);
        Ok(())
    }
}
