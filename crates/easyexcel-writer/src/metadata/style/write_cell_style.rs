//! Mirrors Java `com.alibaba.excel.write.metadata.style.WriteCellStyle`.

use easyexcel_core::ExcelCellStyle;

use crate::metadata::style::write_font::merge_excel_font_style;

/// Mirrors Java `WriteCellStyle`.
///
/// The Java side carries POI-typed fields plus nested `writeFont` and a
/// static `merge` helper. Rust reuses [`ExcelCellStyle`] for the data
/// (including [`ExcelCellStyle::font`]) and mirrors the merge method.
pub type WriteCellStyle = ExcelCellStyle;

/// Mirrors Java `WriteCellStyle.merge(WriteCellStyle source, WriteCellStyle target)`.
///
/// Java merges the source's non-null fields into the target, including
/// nested `WriteFont.merge`. The Rust port performs the same union over
/// [`ExcelCellStyle`]'s `Option` fields and [`ExcelCellStyle::font`].
pub fn merge_write_cell_style(
    source: &ExcelCellStyle,
    mut target: ExcelCellStyle,
) -> ExcelCellStyle {
    macro_rules! or {
        ($field:ident) => {
            if source.$field.is_some() {
                target.$field = source.$field;
            }
        };
    }
    or!(hidden);
    or!(locked);
    or!(quote_prefix);
    or!(horizontal_alignment);
    or!(wrapped);
    or!(vertical_alignment);
    or!(rotation);
    or!(indent);
    or!(border_left);
    or!(border_right);
    or!(border_top);
    or!(border_bottom);
    or!(left_border_color);
    or!(right_border_color);
    or!(top_border_color);
    or!(bottom_border_color);
    or!(fill_pattern);
    or!(fill_background_color);
    or!(fill_foreground_color);
    or!(shrink_to_fit);
    or!(data_format);
    // Java `WriteFont.merge(source.getWriteFont(), target.getWriteFont())`
    if let Some(source_font) = source.font {
        target.font = Some(match target.font {
            Some(existing) => merge_excel_font_style(&source_font, existing),
            None => source_font,
        });
    }
    target
}
