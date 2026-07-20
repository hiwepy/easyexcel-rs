//! Phase 5 — BIFF8 (.xls) encryption placeholder.
//!
//! Java `com.alibaba.excel.support.encrypt.EncryptionInfo` and POI's
//! `HSSFEncryptor` provide standard + agile encryption for legacy
//! XLS workbooks. The Rust port currently returns
//! `ExcelError::Unsupported("password protection is not supported for legacy XLS")`
//! for the `.xls` write/read paths.
//!
//! This stub captures the planned integration points so subsequent
//! phases can fill in:
//!
//! 1. **Standard encryption** (POI compatible, RC4-based)
//!    - `Biff8EncryptionInfo::standard(password) -> EncryptionInfo`
//!    - Wrap with `cfb::CompoundFileWriter` + `Biff8StandardEncryptor`
//!
//! 2. **Agile encryption** (newer OOXML-compatible)
//!    - Already supported via `ms_offcrypto_writer` for XLSX
//!    - For BIFF8: reuse the same crypto primitives but produce an
//!      OLE2 compound document
//!
//! 3. **Decryption read path**
//!    - `OfficeCrypto::decrypt_from_file(...)` for the OOXML container
//!    - For BIFF8: parse `EncryptedPackage` stream + decrypt records
//!
//! Tests in `core_phase5_xls_features_1to1_tests.rs` continue to
//! assert the current `Unsupported` behaviour for these gaps and
//! will be updated when BIFF8 encryption lands.
//!
//! Reference Java files (kept as `Mirrors:` doc anchors):
//! - `com.alibaba.excel.support.encrypt.EncryptionInfo`
//! - `org.apache.poi.hssf.record.crypto.Biff8EncryptionKey`

#![allow(dead_code)]

/// Placeholder marker for the future BIFF8 encryption module.
///
/// When Phase 5.2 lands, this becomes a public re-export of the
/// `Biff8EncryptionInfo` type used by `ExcelWriter::write_with_password`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Biff8EncryptionInfoPlaceholder;

/// Phase 5 entry point marker.
///
/// Mirrors Java `com.alibaba.excel.support.encrypt.EncryptionInfo`
/// subset that will land once BIFF8 standard encryption is implemented.
/// For now this just documents the gap.
pub const PHASE_5_GAP: &str = "Biff8EncryptionInfo (BIFF8 standard encryption) — pending Phase 5.2 implementation";