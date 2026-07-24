//! Mirrors Java `com.alibaba.excel.write.property.ExcelWriteHeadProperty`.

use std::collections::{BTreeMap, HashSet};
use std::ops::Deref;

use crate::metadata::{
    CellRange, ColumnWidthProperty, ConfigurationHolder, ExcelHeadProperty, FontProperty, Head,
    RowHeightProperty, StyleProperty,
};
use crate::{
    ExcelColumn, ExcelError, ExcelFontStyle, ExcelWriteMetadata, HeadKind,
    OnceAbsoluteMergeProperty,
};

/// Mirrors Java `ExcelWriteHeadProperty extends ExcelHeadProperty`.
///
/// This is core metadata rather than a backend object. Keeping it in
/// `easyexcel-core` lets `WriteContextHolder` expose the same resolved property
/// for XLSX, XLS, CSV and template writers without creating a core → writer
/// dependency cycle.
#[derive(Debug, Clone, PartialEq)]
pub struct ExcelWriteHeadProperty {
    inner: ExcelHeadProperty,
    /// Mirrors `ExcelWriteHeadProperty.headRowHeightProperty`.
    pub head_row_height_property: Option<RowHeightProperty>,
    /// Mirrors `ExcelWriteHeadProperty.contentRowHeightProperty`.
    pub content_row_height_property: Option<RowHeightProperty>,
    /// Mirrors `ExcelWriteHeadProperty.onceAbsoluteMergeProperty`.
    pub once_absolute_merge_property: Option<OnceAbsoluteMergeProperty>,
}

impl ExcelWriteHeadProperty {
    /// Creates an empty property.
    #[must_use]
    pub fn new() -> Self {
        Self {
            inner: ExcelHeadProperty::default(),
            head_row_height_property: None,
            content_row_height_property: None,
            once_absolute_merge_property: None,
        }
    }

    /// Resolves a dynamic or class-backed head. (Java constructor)
    #[must_use]
    pub fn from_head(
        configuration_holder: Option<&dyn ConfigurationHolder>,
        head_clazz: Option<String>,
        head: Option<Vec<Vec<String>>>,
        metadata: ExcelWriteMetadata,
    ) -> Self {
        let inner = match head_clazz {
            Some(head_clazz) => {
                ExcelHeadProperty::for_class(configuration_holder, head_clazz, head)
            }
            None => ExcelHeadProperty::new(configuration_holder, head),
        };
        Self::from_inner(inner, metadata)
    }

    /// Resolves Java `Head` entries from Rust derive metadata.
    ///
    /// `columns` contains the effective output column index and schema entry.
    /// An explicit `head` replaces field-derived labels and must contain one
    /// path per effective column.
    ///
    /// # Errors
    ///
    /// Returns [`ExcelError::Format`] when head and column counts differ or a
    /// column index cannot be represented by Java's signed `Integer`.
    pub fn from_columns(
        head_clazz: Option<String>,
        columns: &[(usize, &ExcelColumn)],
        head: Option<&[Vec<String>]>,
        metadata: ExcelWriteMetadata,
    ) -> Result<Self, ExcelError> {
        if let Some(head) = head
            && head.len() != columns.len()
        {
            return Err(ExcelError::Format(format!(
                "head column count {} does not match effective column count {}",
                head.len(),
                columns.len()
            )));
        }

        let mut head_map = BTreeMap::new();
        for (position, (column_index, column)) in columns.iter().enumerate() {
            let column_index = i32::try_from(*column_index).map_err(|_| {
                ExcelError::Format(format!(
                    "head column index {column_index} exceeds Java Integer range"
                ))
            })?;
            let explicit_names = head.and_then(|head| head.get(position));
            let head_names = explicit_names
                .cloned()
                .unwrap_or_else(|| vec![column.name.to_owned()]);
            let mut head_data = Head::new(
                column_index,
                (!column.field.is_empty()).then(|| column.field.to_owned()),
                head_names,
                column.index.is_some(),
                explicit_names.is_some() || !column.name.is_empty(),
            )?;
            head_data.column_width_property = column
                .column_width
                .or(metadata.column_width)
                .map(ColumnWidthProperty::new);
            head_data.loop_merge_property = column.loop_merge;
            head_data.head_style_property = column
                .head_style
                .or(metadata.head_style)
                .map(StyleProperty::new);
            head_data.head_font_property = column
                .head_font_style
                .or(metadata.head_font_style)
                .map(font_property);
            head_map.insert(column_index, head_data);
        }

        let head_kind = if head_clazz.is_some() {
            HeadKind::Class
        } else if head_map.is_empty() {
            HeadKind::None
        } else {
            HeadKind::String
        };
        let inner = ExcelHeadProperty::from_head_map(head_clazz, head_kind, head_map);
        Ok(Self::from_inner(inner, metadata))
    }

    fn from_inner(inner: ExcelHeadProperty, metadata: ExcelWriteMetadata) -> Self {
        Self {
            inner,
            head_row_height_property: metadata.head_row_height.map(RowHeightProperty::new),
            content_row_height_property: metadata.content_row_height.map(RowHeightProperty::new),
            once_absolute_merge_property: metadata.once_absolute_merge,
        }
    }

    /// Returns the underlying inherited header property.
    #[must_use]
    pub const fn inner(&self) -> &ExcelHeadProperty {
        &self.inner
    }

    /// Returns the head row-height property. (Java getter)
    #[must_use]
    pub const fn head_row_height_property(&self) -> Option<&RowHeightProperty> {
        self.head_row_height_property.as_ref()
    }

    /// Returns the content row-height property. (Java getter)
    #[must_use]
    pub const fn content_row_height_property(&self) -> Option<&RowHeightProperty> {
        self.content_row_height_property.as_ref()
    }

    /// Returns the once-absolute merge property. (Java getter)
    #[must_use]
    pub const fn once_absolute_merge_property(&self) -> Option<&OnceAbsoluteMergeProperty> {
        self.once_absolute_merge_property.as_ref()
    }

    /// Calculates every automatic header merge. (Java `headCellRangeList()`)
    #[must_use]
    pub fn head_cell_range_list(&self) -> Vec<CellRange> {
        let head_list = self.inner.head_map.values().collect::<Vec<_>>();
        let mut already_ranged = HashSet::new();
        let mut ranges = Vec::new();

        for (column_position, head) in head_list.iter().enumerate() {
            for row in 0..head.head_name_list.len() {
                if !already_ranged.insert((column_position, row)) {
                    continue;
                }
                let name = &head.head_name_list[row];
                let mut last_column_position = column_position;
                let mut last_row = row;
                for candidate in column_position + 1..head_list.len() {
                    let key = (candidate, row);
                    if head_list[candidate].head_name_list[row] == *name
                        && already_ranged.insert(key)
                    {
                        last_column_position = candidate;
                    } else {
                        break;
                    }
                }

                'rows: for candidate_row in row + 1..head.head_name_list.len() {
                    let mut row_cells = Vec::new();
                    for candidate_column in column_position..=last_column_position {
                        let key = (candidate_column, candidate_row);
                        if head_list[candidate_column].head_name_list[candidate_row] != *name
                            || already_ranged.contains(&key)
                        {
                            break 'rows;
                        }
                        row_cells.push(key);
                    }
                    already_ranged.extend(row_cells);
                    last_row = candidate_row;
                }

                if row == last_row && column_position == last_column_position {
                    continue;
                }
                let first_col = head.column_index.unwrap_or(column_position as i32);
                let last_col = head_list[last_column_position]
                    .column_index
                    .unwrap_or(last_column_position as i32);
                ranges.push(CellRange::new(
                    row as i32,
                    last_row as i32,
                    first_col,
                    last_col,
                ));
            }
        }
        ranges
    }
}

impl Default for ExcelWriteHeadProperty {
    fn default() -> Self {
        Self::new()
    }
}

impl Deref for ExcelWriteHeadProperty {
    type Target = ExcelHeadProperty;

    fn deref(&self) -> &Self::Target {
        &self.inner
    }
}

fn font_property(font: ExcelFontStyle) -> FontProperty {
    FontProperty {
        font_name: font.font_name,
        font_height_in_points: font.font_height_in_points,
        italic: font.italic,
        strikeout: font.strikeout,
        color: font.color,
        type_offset: font.type_offset,
        underline: font.underline,
        charset: font.charset,
        bold: font.bold,
    }
}
