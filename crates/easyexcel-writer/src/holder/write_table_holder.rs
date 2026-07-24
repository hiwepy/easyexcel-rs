//! Mirrors Java `com.alibaba.excel.write.metadata.holder.WriteTableHolder`.

use std::ops::{Deref, DerefMut};

use crate::holder::abstract_write_holder::AbstractWriteHolder;
use crate::metadata::WriteBasicParameter;

/// Mirrors Java `WriteTableHolder extends AbstractWriteHolder`.
///
/// Java's holder carries a POI `Sheet` plus a `tableNo` field. The Rust port
/// mirrors the type so the [`crate::ExcelWriterTableBuilder`] can return a
/// `WriteTableHolder` for parity. Runtime callbacks expose the active table
/// through [`easyexcel_core::WriteTableHolderView`].
pub struct WriteTableHolder<'a> {
    abstract_holder: AbstractWriteHolder,
    table_no: i32,
    parent_sheet: Option<&'a str>,
    last_row_index: i32,
}

impl<'a> WriteTableHolder<'a> {
    /// Creates a table holder matching the Java `WriteTableHolder(WriteTable, WriteSheetHolder)` initialiser.
    #[must_use]
    pub fn new(table_no: i32) -> Self {
        Self {
            abstract_holder: AbstractWriteHolder::default(),
            table_no,
            parent_sheet: None,
            last_row_index: 0,
        }
    }

    /// Creates a table holder and resolves nullable values against its sheet.
    #[must_use]
    pub fn from_parameter(
        table_no: i32,
        parameter: &WriteBasicParameter,
        parent: &AbstractWriteHolder,
    ) -> Self {
        let mut holder = Self::new(table_no);
        holder.abstract_holder = AbstractWriteHolder::from_parameter(parameter, Some(parent));
        holder
    }

    /// Returns the inherited write-holder state.
    #[must_use]
    pub const fn abstract_holder(&self) -> &AbstractWriteHolder {
        &self.abstract_holder
    }

    /// Returns mutable inherited write-holder state.
    pub const fn abstract_holder_mut(&mut self) -> &mut AbstractWriteHolder {
        &mut self.abstract_holder
    }

    /// Returns the parent sheet name, if any. (Java `getParentWriteSheetHolder().getSheetName()`)
    #[must_use]
    pub fn parent_sheet(&self) -> Option<&str> {
        self.parent_sheet
    }

    /// Sets the parent sheet name.
    pub fn set_parent_sheet(&mut self, parent: &'a str) {
        self.parent_sheet = Some(parent);
    }

    /// Returns the zero-based table index. (Java `getTableNo()`)
    #[must_use]
    pub const fn table_no(&self) -> i32 {
        self.table_no
    }

    /// Returns the last row index. (Java `getLastRowIndex()`)
    #[must_use]
    pub const fn last_row_index(&self) -> i32 {
        self.last_row_index
    }
}

impl Deref for WriteTableHolder<'_> {
    type Target = AbstractWriteHolder;

    fn deref(&self) -> &Self::Target {
        &self.abstract_holder
    }
}

impl DerefMut for WriteTableHolder<'_> {
    fn deref_mut(&mut self) -> &mut Self::Target {
        &mut self.abstract_holder
    }
}

#[cfg(test)]
mod tests {
    use std::collections::HashSet;

    use super::*;
    use crate::WriteHolder;
    use crate::holder::write_sheet_holder::WriteSheetHolder;
    use crate::holder::write_workbook_holder::WriteWorkbookHolder;

    #[test]
    fn workbook_sheet_table_holders_resolve_java_parent_chain() {
        let workbook = WriteWorkbookHolder::from_parameter(
            "out.xlsx",
            &WriteBasicParameter {
                need_head: Some(false),
                include_column_indexes: Some(vec![1, 3]),
                ..WriteBasicParameter::default()
            },
        );
        let sheet = WriteSheetHolder::from_parameter(
            "Data",
            0,
            &WriteBasicParameter {
                need_head: Some(true),
                exclude_column_field_names: Some(vec!["secret".to_owned()]),
                ..WriteBasicParameter::default()
            },
            workbook.abstract_holder(),
        );
        let table = WriteTableHolder::from_parameter(
            2,
            &WriteBasicParameter {
                include_column_indexes: Some(Vec::new()),
                order_by_include_column: Some(true),
                ..WriteBasicParameter::default()
            },
            sheet.abstract_holder(),
        );

        assert!(!workbook.need_head());
        assert!(sheet.need_head());
        assert!(table.need_head());
        assert_eq!(sheet.include_column_indexes, Some(HashSet::from([1, 3])));
        assert_eq!(table.include_column_indexes, Some(HashSet::new()));
        assert_eq!(
            table.exclude_column_field_names,
            Some(HashSet::from(["secret".to_owned()]))
        );
        assert!(table.order_by_include_column());
    }
}
