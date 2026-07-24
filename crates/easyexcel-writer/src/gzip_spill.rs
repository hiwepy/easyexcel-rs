//! SXSSF `GZIPSheetDataWriter` equivalent: gzip-compressed row spill on disk.
//!
//! Java `SXSSFWorkbook.setCompressTempFiles(true)` routes sheet XML through
//! `GZIPSheetDataWriter`. `rust_xlsxwriter` constant-memory mode cannot gzip its
//! internal tempfile, so this module owns the durable spill while rows are
//! written; [`ExcelWriter`] materializes into a constant-memory worksheet only
//! at `finish` (stream decode → write → ZIP), keeping peak RAM bounded.

use std::fs::{File, OpenOptions};
use std::io::{Read, Write};
use std::path::{Path, PathBuf};

use bigdecimal::BigDecimal;
use chrono::{NaiveDate, NaiveDateTime};
use easyexcel_core::{CellValue, ExcelError, ImageData, Result, RichTextStringData};
use flate2::Compression;
use flate2::read::GzDecoder;
use flate2::write::GzEncoder;
use tempfile::{Builder, TempDir};

/// Gzip magic number (`1f 8b`) — used by tests to observe true compression.
pub const GZIP_MAGIC: [u8; 2] = [0x1f, 0x8b];

/// Observable snapshot of an active or finished gzip spill file.
#[derive(Debug, Clone)]
pub struct GzipSpillSnapshot {
    /// Logical sheet name this spill belongs to.
    pub sheet_name: String,
    /// Path of the gzip tempfile (named, so tests can open it).
    pub path: PathBuf,
    /// Whether the file begins with gzip magic.
    pub is_gzip: bool,
    /// On-disk compressed size in bytes.
    pub compressed_len: u64,
    /// Uncompressed payload bytes written into the encoder.
    pub uncompressed_len: u64,
}

/// Streaming gzip spill writer mirroring POI `GZIPSheetDataWriter`.
pub struct GzipSheetDataWriter {
    sheet_name: String,
    path: PathBuf,
    encoder: GzEncoder<File>,
    uncompressed_len: u64,
    /// Keeps the parent temp directory alive for the spill lifetime.
    _dir: Option<TempDir>,
}

impl GzipSheetDataWriter {
    /// Creates a new gzip spill file under `dir` for `sheet_name`.
    ///
    /// # Errors
    ///
    /// Returns an I/O error when the tempfile cannot be created.
    pub fn create(dir: &Path, sheet_name: impl Into<String>) -> Result<Self> {
        let sheet_name = sheet_name.into();
        let tmp = Builder::new()
            .prefix("easyexcel-sxssf-")
            .suffix(".xml.gz")
            .tempfile_in(dir)
            .map_err(ExcelError::Io)?;
        // tempfile 3.x: `NamedTempFile::keep` → `(File, PathBuf)`.
        let (file, path) = tmp.keep().map_err(|error| ExcelError::Io(error.error))?;
        let encoder = GzEncoder::new(file, Compression::default());
        Ok(Self {
            sheet_name,
            path,
            encoder,
            uncompressed_len: 0,
            _dir: None,
        })
    }

    /// Creates a spill that owns its temporary directory (deleted on drop).
    ///
    /// # Errors
    ///
    /// Returns an I/O error when the temp directory or file cannot be created.
    pub fn create_owned(sheet_name: impl Into<String>) -> Result<Self> {
        let dir = TempDir::new().map_err(ExcelError::Io)?;
        let mut writer = Self::create(dir.path(), sheet_name)?;
        writer._dir = Some(dir);
        Ok(writer)
    }

    /// Appends one data row (cell values) to the gzip spill.
    ///
    /// # Errors
    ///
    /// Returns a format or I/O error when encoding or writing fails.
    pub fn write_row(&mut self, cells: &[CellValue]) -> Result<()> {
        let mut buf = Vec::with_capacity(64 + cells.len() * 16);
        encode_row(&mut buf, cells)?;
        self.encoder.write_all(&buf).map_err(ExcelError::Io)?;
        self.uncompressed_len = self
            .uncompressed_len
            .saturating_add(u64::try_from(buf.len()).unwrap_or(u64::MAX));
        Ok(())
    }

    /// Flushes buffered gzip bytes so magic / size are observable on disk.
    ///
    /// # Errors
    ///
    /// Returns an I/O error on flush failure.
    pub fn flush(&mut self) -> Result<()> {
        self.encoder.flush().map_err(ExcelError::Io)
    }

    /// Returns a snapshot suitable for tests (gzip magic + sizes).
    ///
    /// # Errors
    ///
    /// Returns an I/O error when flushing or stating the file fails.
    pub fn snapshot(&mut self) -> Result<GzipSpillSnapshot> {
        self.flush()?;
        let compressed_len = std::fs::metadata(&self.path)
            .map(|meta| meta.len())
            .unwrap_or(0);
        let is_gzip = file_has_gzip_magic(&self.path);
        Ok(GzipSpillSnapshot {
            sheet_name: self.sheet_name.clone(),
            path: self.path.clone(),
            is_gzip,
            compressed_len,
            uncompressed_len: self.uncompressed_len,
        })
    }

    /// Finishes the encoder and returns a readable spill handle.
    ///
    /// # Errors
    ///
    /// Returns an I/O error when finishing gzip or reopening the file fails.
    pub fn finish(self) -> Result<GzipSpillReader> {
        let path = self.path;
        let uncompressed_len = self.uncompressed_len;
        let sheet_name = self.sheet_name;
        let _dir = self._dir;
        self.encoder.finish().map_err(ExcelError::Io)?;
        let compressed_len = std::fs::metadata(&path).map(|meta| meta.len()).unwrap_or(0);
        let file = OpenOptions::new()
            .read(true)
            .open(&path)
            .map_err(ExcelError::Io)?;
        Ok(GzipSpillReader {
            sheet_name,
            path,
            decoder: GzDecoder::new(file),
            uncompressed_len,
            compressed_len,
            _dir,
        })
    }
}

/// Read side of a finished gzip spill (stream decode, constant memory).
pub struct GzipSpillReader {
    sheet_name: String,
    path: PathBuf,
    decoder: GzDecoder<File>,
    uncompressed_len: u64,
    compressed_len: u64,
    _dir: Option<TempDir>,
}

impl GzipSpillReader {
    /// Returns spill metadata after finish.
    #[must_use]
    pub fn snapshot(&self) -> GzipSpillSnapshot {
        GzipSpillSnapshot {
            sheet_name: self.sheet_name.clone(),
            path: self.path.clone(),
            is_gzip: file_has_gzip_magic(&self.path),
            compressed_len: self.compressed_len,
            uncompressed_len: self.uncompressed_len,
        }
    }

    /// Decodes the next spilled row, or `None` at EOF.
    ///
    /// # Errors
    ///
    /// Returns a format or I/O error when the stream is corrupt.
    pub fn next_row(&mut self) -> Result<Option<Vec<CellValue>>> {
        let mut len_buf = [0u8; 4];
        match self.decoder.read_exact(&mut len_buf) {
            Ok(()) => {}
            Err(error) if error.kind() == std::io::ErrorKind::UnexpectedEof => return Ok(None),
            Err(error) => return Err(ExcelError::Io(error)),
        }
        let len = u32::from_le_bytes(len_buf) as usize;
        let mut payload = vec![0u8; len];
        self.decoder
            .read_exact(&mut payload)
            .map_err(ExcelError::Io)?;
        let cells = decode_row(&payload)?;
        Ok(Some(cells))
    }
}

/// Returns whether `path` starts with gzip magic bytes.
#[must_use]
pub fn file_has_gzip_magic(path: &Path) -> bool {
    let Ok(mut file) = File::open(path) else {
        return false;
    };
    let mut magic = [0u8; 2];
    matches!(file.read_exact(&mut magic), Ok(())) && magic == GZIP_MAGIC
}

fn encode_row(out: &mut Vec<u8>, cells: &[CellValue]) -> Result<()> {
    let mut body = Vec::with_capacity(cells.len() * 16);
    write_u32(
        &mut body,
        u32::try_from(cells.len())
            .map_err(|_| ExcelError::Format("row cell count exceeds u32".to_owned()))?,
    );
    for cell in cells {
        encode_cell(&mut body, cell)?;
    }
    write_u32(
        out,
        u32::try_from(body.len())
            .map_err(|_| ExcelError::Format("spill row exceeds u32 length".to_owned()))?,
    );
    out.extend_from_slice(&body);
    Ok(())
}

fn decode_row(payload: &[u8]) -> Result<Vec<CellValue>> {
    let mut cursor = 0usize;
    let count = read_u32(payload, &mut cursor)? as usize;
    let mut cells = Vec::with_capacity(count);
    for _ in 0..count {
        cells.push(decode_cell(payload, &mut cursor)?);
    }
    Ok(cells)
}

fn encode_cell(out: &mut Vec<u8>, value: &CellValue) -> Result<()> {
    match value {
        CellValue::Empty => out.push(0),
        CellValue::String(text) => {
            out.push(1);
            write_str(out, text)?;
        }
        CellValue::Bool(flag) => {
            out.push(2);
            out.push(u8::from(*flag));
        }
        CellValue::Int(number) => {
            out.push(3);
            out.extend_from_slice(&number.to_le_bytes());
        }
        CellValue::Float(number) => {
            out.push(4);
            out.extend_from_slice(&number.to_le_bytes());
        }
        CellValue::Decimal(number) => {
            out.push(5);
            write_str(out, &number.to_string())?;
        }
        CellValue::Date(date) => {
            out.push(6);
            write_str(out, &date.format("%Y-%m-%d").to_string())?;
        }
        CellValue::DateTime(date_time) => {
            out.push(7);
            write_str(out, &date_time.format("%Y-%m-%d %H:%M:%S%.f").to_string())?;
        }
        CellValue::Error(text) => {
            out.push(8);
            write_str(out, text)?;
        }
        CellValue::Formula(text) => {
            out.push(9);
            write_str(out, text)?;
        }
        CellValue::Hyperlink { url, text } => {
            out.push(10);
            write_str(out, url)?;
            write_str(out, text)?;
        }
        CellValue::Comment { value, text } => {
            out.push(11);
            write_str(out, text)?;
            encode_cell(out, value)?;
        }
        CellValue::Image(bytes) => {
            out.push(12);
            write_bytes(out, bytes)?;
        }
        CellValue::RichText(rich) => {
            // Fonts are not required for compress-temp spill round-trips.
            out.push(13);
            write_str(out, rich.text_string())?;
        }
        CellValue::Images { value, images } => {
            out.push(14);
            encode_cell(out, value)?;
            write_u32(
                out,
                u32::try_from(images.len())
                    .map_err(|_| ExcelError::Format("image list exceeds u32".to_owned()))?,
            );
            for image in images {
                write_bytes(out, image.image())?;
            }
        }
    }
    Ok(())
}

fn decode_cell(buf: &[u8], cursor: &mut usize) -> Result<CellValue> {
    let tag = read_u8(buf, cursor)?;
    Ok(match tag {
        0 => CellValue::Empty,
        1 => CellValue::String(read_str(buf, cursor)?),
        2 => CellValue::Bool(read_u8(buf, cursor)? != 0),
        3 => {
            let bytes = read_exact::<8>(buf, cursor)?;
            CellValue::Int(i64::from_le_bytes(bytes))
        }
        4 => {
            let bytes = read_exact::<8>(buf, cursor)?;
            CellValue::Float(f64::from_le_bytes(bytes))
        }
        5 => {
            let text = read_str(buf, cursor)?;
            let number: BigDecimal = text
                .parse()
                .map_err(|error| ExcelError::Format(format!("invalid decimal spill: {error}")))?;
            CellValue::Decimal(number)
        }
        6 => {
            let text = read_str(buf, cursor)?;
            let date = NaiveDate::parse_from_str(&text, "%Y-%m-%d")
                .map_err(|error| ExcelError::Format(format!("invalid date spill: {error}")))?;
            CellValue::Date(date)
        }
        7 => {
            let text = read_str(buf, cursor)?;
            let date_time = NaiveDateTime::parse_from_str(&text, "%Y-%m-%d %H:%M:%S%.f")
                .or_else(|_| NaiveDateTime::parse_from_str(&text, "%Y-%m-%d %H:%M:%S"))
                .map_err(|error| ExcelError::Format(format!("invalid datetime spill: {error}")))?;
            CellValue::DateTime(date_time)
        }
        8 => CellValue::Error(read_str(buf, cursor)?),
        9 => CellValue::Formula(read_str(buf, cursor)?),
        10 => CellValue::Hyperlink {
            url: read_str(buf, cursor)?,
            text: read_str(buf, cursor)?,
        },
        11 => {
            let text = read_str(buf, cursor)?;
            let value = Box::new(decode_cell(buf, cursor)?);
            CellValue::Comment { value, text }
        }
        12 => CellValue::Image(read_bytes(buf, cursor)?),
        13 => CellValue::RichText(RichTextStringData::new(read_str(buf, cursor)?)),
        14 => {
            let value = Box::new(decode_cell(buf, cursor)?);
            let count = read_u32(buf, cursor)? as usize;
            let mut images = Vec::with_capacity(count);
            for _ in 0..count {
                images.push(ImageData::new(read_bytes(buf, cursor)?));
            }
            CellValue::Images { value, images }
        }
        other => {
            return Err(ExcelError::Format(format!(
                "unknown gzip spill cell tag: {other}"
            )));
        }
    })
}

fn write_u32(out: &mut Vec<u8>, value: u32) {
    out.extend_from_slice(&value.to_le_bytes());
}

fn write_str(out: &mut Vec<u8>, value: &str) -> Result<()> {
    write_bytes(out, value.as_bytes())
}

fn write_bytes(out: &mut Vec<u8>, value: &[u8]) -> Result<()> {
    write_u32(
        out,
        u32::try_from(value.len())
            .map_err(|_| ExcelError::Format("spill byte length exceeds u32".to_owned()))?,
    );
    out.extend_from_slice(value);
    Ok(())
}

fn read_u8(buf: &[u8], cursor: &mut usize) -> Result<u8> {
    let value = *buf
        .get(*cursor)
        .ok_or_else(|| ExcelError::Format("gzip spill truncated (u8)".to_owned()))?;
    *cursor += 1;
    Ok(value)
}

fn read_u32(buf: &[u8], cursor: &mut usize) -> Result<u32> {
    let bytes = read_exact::<4>(buf, cursor)?;
    Ok(u32::from_le_bytes(bytes))
}

fn read_exact<const N: usize>(buf: &[u8], cursor: &mut usize) -> Result<[u8; N]> {
    let end = cursor
        .checked_add(N)
        .ok_or_else(|| ExcelError::Format("gzip spill cursor overflow".to_owned()))?;
    let slice = buf
        .get(*cursor..end)
        .ok_or_else(|| ExcelError::Format("gzip spill truncated".to_owned()))?;
    let mut out = [0u8; N];
    out.copy_from_slice(slice);
    *cursor = end;
    Ok(out)
}

fn read_str(buf: &[u8], cursor: &mut usize) -> Result<String> {
    let bytes = read_bytes(buf, cursor)?;
    String::from_utf8(bytes)
        .map_err(|error| ExcelError::Format(format!("gzip spill utf-8: {error}")))
}

fn read_bytes(buf: &[u8], cursor: &mut usize) -> Result<Vec<u8>> {
    let len = read_u32(buf, cursor)? as usize;
    let end = cursor
        .checked_add(len)
        .ok_or_else(|| ExcelError::Format("gzip spill cursor overflow".to_owned()))?;
    let slice = buf
        .get(*cursor..end)
        .ok_or_else(|| ExcelError::Format("gzip spill truncated (bytes)".to_owned()))?;
    *cursor = end;
    Ok(slice.to_vec())
}

#[cfg(test)]
mod tests {
    use super::*;
    use chrono::NaiveDate;

    #[test]
    fn gzip_spill_round_trips_cells_and_exposes_magic() {
        let mut writer = GzipSheetDataWriter::create_owned("Sheet1").expect("create");
        let date = NaiveDate::from_ymd_opt(2020, 1, 1).expect("date");
        writer
            .write_row(&[
                CellValue::String("字符串0".to_owned()),
                CellValue::Date(date),
                CellValue::Float(0.56),
                CellValue::Int(42),
                CellValue::Bool(true),
                CellValue::Empty,
            ])
            .expect("write");
        let snap = writer.snapshot().expect("snapshot");
        assert!(snap.is_gzip, "spill must start with gzip magic");
        assert!(snap.uncompressed_len > 0);
        assert!(snap.compressed_len > 0);
        // Highly repetitive / small payloads may not shrink, but magic must be present.
        assert_eq!(&snap.path.extension().and_then(|e| e.to_str()), &Some("gz"));

        let mut reader = writer.finish().expect("finish");
        let row = reader.next_row().expect("decode").expect("one row");
        assert_eq!(row[0], CellValue::String("字符串0".to_owned()));
        assert_eq!(row[1], CellValue::Date(date));
        assert!(matches!(row[2], CellValue::Float(v) if (v - 0.56).abs() < f64::EPSILON));
        assert_eq!(row[3], CellValue::Int(42));
        assert_eq!(row[4], CellValue::Bool(true));
        assert_eq!(row[5], CellValue::Empty);
        assert!(reader.next_row().expect("eof").is_none());
        assert!(reader.snapshot().is_gzip);
    }
}
