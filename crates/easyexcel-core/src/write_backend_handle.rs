//! Backend-neutral row/cell handles exposed to write handlers.

use std::cell::RefCell;

use crate::{CellValue, ExcelCellStyle};

/// Backend-neutral equivalent of POI's mutable `Cell` callback object.
///
/// Mutations are recorded and committed by the active writer backend after
/// the logical callback chain. This never pretends to be an Apache POI cell.
#[derive(Debug, Clone, PartialEq)]
pub struct WriteCellHandle {
    row_index: u32,
    column_index: u16,
    current_value: RefCell<CellValue>,
    requested_value: RefCell<Option<CellValue>>,
    requested_style: RefCell<Option<ExcelCellStyle>>,
    requested_skip: RefCell<Option<bool>>,
}

impl WriteCellHandle {
    /// Creates a handle for one physical cell.
    #[must_use]
    pub fn new(row_index: u32, column_index: u16, initial_value: CellValue) -> Self {
        Self {
            row_index,
            column_index,
            current_value: RefCell::new(initial_value),
            requested_value: RefCell::new(None),
            requested_style: RefCell::new(None),
            requested_skip: RefCell::new(None),
        }
    }

    /// Returns the zero-based physical row.
    #[must_use]
    pub const fn row_index(&self) -> u32 {
        self.row_index
    }

    /// Returns the zero-based physical column.
    #[must_use]
    pub const fn column_index(&self) -> u16 {
        self.column_index
    }

    /// Returns the latest logical value visible to the callback chain.
    #[must_use]
    pub fn value(&self) -> CellValue {
        self.current_value.borrow().clone()
    }

    /// Requests a final cell value, including from `afterCellDispose`.
    pub fn set_value(&self, value: CellValue) {
        *self.current_value.borrow_mut() = value.clone();
        *self.requested_value.borrow_mut() = Some(value);
    }

    /// Synchronizes a value changed through the compatibility context field.
    pub fn sync_value(&self, value: &CellValue) {
        *self.current_value.borrow_mut() = value.clone();
    }

    /// Requests a final backend-neutral cell style.
    pub fn set_style(&self, style: ExcelCellStyle) {
        *self.requested_style.borrow_mut() = Some(style);
    }

    /// Requests that the physical cell be omitted or restored.
    pub fn set_skipped(&self, skipped: bool) {
        *self.requested_skip.borrow_mut() = Some(skipped);
    }

    /// Returns the requested value override.
    #[must_use]
    pub fn requested_value(&self) -> Option<CellValue> {
        self.requested_value.borrow().clone()
    }

    /// Returns the requested style override.
    #[must_use]
    pub fn requested_style(&self) -> Option<ExcelCellStyle> {
        *self.requested_style.borrow()
    }

    /// Returns the requested skip override.
    #[must_use]
    pub fn requested_skip(&self) -> Option<bool> {
        *self.requested_skip.borrow()
    }
}

/// Backend-neutral equivalent of POI's mutable `Row` callback object.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WriteRowHandle {
    row_index: u32,
    requested_height: RefCell<Option<u16>>,
}

impl WriteRowHandle {
    /// Creates a handle for one physical row.
    #[must_use]
    pub fn new(row_index: u32) -> Self {
        Self {
            row_index,
            requested_height: RefCell::new(None),
        }
    }

    /// Returns the zero-based physical row.
    #[must_use]
    pub const fn row_index(&self) -> u32 {
        self.row_index
    }

    /// Requests a final row height in points.
    pub fn set_height(&self, height: u16) {
        *self.requested_height.borrow_mut() = Some(height);
    }

    /// Returns the requested final row height.
    #[must_use]
    pub fn requested_height(&self) -> Option<u16> {
        *self.requested_height.borrow()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn logical_handles_record_mutations_without_fake_backend_objects() {
        let row = WriteRowHandle::new(4);
        row.set_height(27);
        assert_eq!(row.row_index(), 4);
        assert_eq!(row.requested_height(), Some(27));

        let cell = WriteCellHandle::new(4, 2, CellValue::String("source".to_owned()));
        cell.set_value(CellValue::String("changed".to_owned()));
        cell.set_style(ExcelCellStyle {
            hidden: Some(true),
            ..ExcelCellStyle::new()
        });
        cell.set_skipped(false);
        assert_eq!(cell.row_index(), 4);
        assert_eq!(cell.column_index(), 2);
        assert_eq!(
            cell.requested_value(),
            Some(CellValue::String("changed".to_owned()))
        );
        assert_eq!(cell.value(), CellValue::String("changed".to_owned()));
        assert_eq!(
            cell.requested_style().and_then(|style| style.hidden),
            Some(true)
        );
        assert_eq!(cell.requested_skip(), Some(false));
    }
}
