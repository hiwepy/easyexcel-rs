//! Mirrors Java `com.alibaba.excel.metadata.Head`.

use crate::excel_error::ExcelError;
use crate::metadata::property::{ColumnWidthProperty, FontProperty, LoopMergeProperty, StyleProperty};

/// Excel header metadata for one column.
///
/// Rust port of Java `Head`.
#[derive(Debug, Clone, PartialEq)]
pub struct Head {
    /// Column index. (Java `columnIndex`)
    pub column_index: Option<i32>,
    /// Rust field name when bound to a model class. (Java `fieldName`)
    pub field_name: Option<String>,
    /// Header labels from the top row down. (Java `headNameList`)
    pub head_name_list: Vec<String>,
    /// Whether `@ExcelProperty.index` forced the column index. (Java `forceIndex`)
    pub force_index: bool,
    /// Whether `@ExcelProperty.value` forced the header name. (Java `forceName`)
    pub force_name: bool,
    /// Column width annotation. (Java `columnWidthProperty`)
    pub column_width_property: Option<ColumnWidthProperty>,
    /// Loop merge annotation. (Java `loopMergeProperty`)
    pub loop_merge_property: Option<LoopMergeProperty>,
    /// Header style annotation. (Java `headStyleProperty`)
    pub head_style_property: Option<StyleProperty>,
    /// Header font annotation. (Java `headFontProperty`)
    pub head_font_property: Option<FontProperty>,
}

impl Head {
    /// Creates a head definition. (Java constructor)
    ///
    /// # Errors
    ///
    /// Returns [`ExcelError::Format`] when any header label is null/empty in
    /// the Java sense (Rust rejects empty strings in the name list).
    pub fn new(
        column_index: i32,
        field_name: Option<String>,
        head_name_list: Vec<String>,
        force_index: bool,
        force_name: bool,
    ) -> Result<Self, ExcelError> {
        for head_name in &head_name_list {
            if head_name.is_empty() {
                return Err(ExcelError::Format(
                    "head name can not be null.".to_owned(),
                ));
            }
        }

        Ok(Self {
            column_index: Some(column_index),
            field_name,
            head_name_list,
            force_index,
            force_name,
            column_width_property: None,
            loop_merge_property: None,
            head_style_property: None,
            head_font_property: None,
        })
    }

    /// Returns the column index. (Java `getColumnIndex()`)
    #[must_use]
    pub fn column_index(&self) -> Option<i32> {
        self.column_index
    }

    /// Returns the field name. (Java `getFieldName()`)
    #[must_use]
    pub fn field_name(&self) -> Option<&str> {
        self.field_name.as_deref()
    }

    /// Returns the header labels. (Java `getHeadNameList()`)
    #[must_use]
    pub fn head_name_list(&self) -> &[String] {
        &self.head_name_list
    }

    /// Returns whether the column index was forced. (Java `getForceIndex()`)
    #[must_use]
    pub const fn force_index(&self) -> bool {
        self.force_index
    }

    /// Returns whether the header name was forced. (Java `getForceName()`)
    #[must_use]
    pub const fn force_name(&self) -> bool {
        self.force_name
    }
}
