//! Minimal BIFF8 (Excel 97–2003 `.xls`) writer.
//!
//! # Java mapping
//!
//! | Java EasyExcel | Rust |
//! |---|---|
//! | `excelType(ExcelTypeEnum.XLS)` | path / stream ending in `.xls` |
//! | `EasyExcel.write(...).sheet().doWrite(data)` | [`crate::write_xls`] / `ExcelWriterBuilder::do_write` |
//! | `ExcelWriter.write(data, writeSheet)` | [`crate::ExcelWriter::write`] on a `.xls` path |
//! | POI `HSSFWorkbook` | [`Biff8Book`] + OLE/CFB `Workbook` stream |
//! | `sheet.setColumnWidth` / `@ColumnWidth` | [`Biff8Sheet::set_column_width`] → COLINFO |
//! | `row.setHeightInPoints` / `@HeadRowHeight` | [`Biff8Sheet::set_row_height`] → ROW |
//! | `WriteCellStyle` / `WriteFont` / `IndexedColors` | [`style::Biff8StyleTable`] → FONT + XF |
//! | `addMergedRegion` / `@ContentLoopMerge` | [`Biff8Sheet::add_merge`] → MERGECELLS |
//!
//! # Capability boundary (deliberately minimal)
//!
//! **Supported:** strings (SST), numbers, booleans, dates/datetimes (1900 system),
//! single or multiple sheets, header row + data rows, column widths, row heights,
//! basic fonts (bold/italic/size/indexed colour), solid fill colours, merge
//! regions, calamine / `EasyExcel::read` round-trip for those scalars + merges.
//!
//! **Still unsupported (explicit `Unsupported` or degraded):**
//! - **Password / RC4 / XOR encryption** — typed `Unsupported` (not OOXML Agile).
//! - **Images** — typed `Unsupported` for `CellValue::Image` / non-empty `Images`
//!   (no MSODrawing/OBJ/Escher; never silently drop image bytes).
//! - in-place OLE style/merge preservation beyond template MVP, collection/horizontal
//!   `.xls` fill, true formula tokens, hyperlink/comment records, rich-text runs,
//!   borders, arbitrary custom number formats, charts, macros.
//!
//! Gaps fail visibly — never silently rewrite as XLSX.
//!
//! **Template append (record-preserving MVP):** see [`template`] —
//! [`Biff8TemplatePackage`] keeps unmodified BIFF records and appends/overwrites
//! cells on an existing sheet. Placeholder `{key}` fill and brand-new sheets remain
//! unsupported at higher layers.

mod encode;
pub mod encrypt;
mod style;
mod template;
mod workbook;

pub use workbook::{
    Biff8Book, Biff8Cell, Biff8Merge, Biff8Sheet, Biff8Value, date_to_excel_serial,
    date_to_excel_serial_with_windowing, datetime_to_excel_serial,
    datetime_to_excel_serial_with_windowing,
};

pub use encrypt::{Biff8EncryptionInfoPlaceholder, PHASE_5_GAP};

pub use encode::{XF_DATE, XF_DATETIME, XF_GENERAL};
pub use style::{Biff8StyleRequest, Biff8StyleTable};
pub use template::{Biff8TemplatePackage, looks_like_xls};
