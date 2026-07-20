//! Mirrors Java `com.alibaba.excel.write.metadata.style.WriteFont`.

use easyexcel_core::{ExcelFontStyle, WriteFont};

/// Mirrors Java `WriteCellStyle`'s font side.
pub type WriteCellFont = WriteFont;

/// Mirrors Java `WriteFont.merge(WriteFont source, WriteFont target)`.
///
/// Java's `merge` copies every non-`None` field from source to target.
/// The Rust port performs the same union over the `Option` fields on
/// [`WriteFont`].
pub fn merge_write_font(source: &WriteFont, mut target: WriteFont) -> WriteFont {
    if source.get_font_name().is_some() {
        target = target.font_name(source.get_font_name().unwrap().to_owned());
    }
    if let Some(height) = source.get_font_height_in_points() {
        target = target.font_height_in_points(height);
    }
    if let Some(italic) = source.get_italic() {
        target = target.italic(italic);
    }
    if let Some(strikeout) = source.get_strikeout() {
        target = target.strikeout(strikeout);
    }
    if let Some(color) = source.get_color() {
        target = target.color(color);
    }
    if let Some(script) = source.get_type_offset() {
        target = target.type_offset(script);
    }
    if let Some(underline) = source.get_underline() {
        target = target.underline(underline);
    }
    if let Some(charset) = source.get_charset() {
        target = target.charset(charset);
    }
    if let Some(bold) = source.get_bold() {
        target = target.bold(bold);
    }
    target
}

/// Merges annotation/strategy fonts. (Java `WriteFont.merge` over `ExcelFontStyle`)
///
/// Copies every non-`None` field from `source` onto `target`, matching the
/// Java `WriteFont.merge` null-skip semantics used when nesting fonts inside
/// `WriteCellStyle`.
#[must_use]
pub fn merge_excel_font_style(
    source: &ExcelFontStyle,
    mut target: ExcelFontStyle,
) -> ExcelFontStyle {
    if source.font_name.is_some() {
        target.font_name = source.font_name;
    }
    if source.font_height_in_points.is_some() {
        target.font_height_in_points = source.font_height_in_points;
    }
    if source.italic.is_some() {
        target.italic = source.italic;
    }
    if source.strikeout.is_some() {
        target.strikeout = source.strikeout;
    }
    if source.color.is_some() {
        target.color = source.color;
    }
    if source.type_offset.is_some() {
        target.type_offset = source.type_offset;
    }
    if source.underline.is_some() {
        target.underline = source.underline;
    }
    if source.charset.is_some() {
        target.charset = source.charset;
    }
    if source.bold.is_some() {
        target.bold = source.bold;
    }
    target
}

/// Converts runtime [`WriteFont`] into Copy [`ExcelFontStyle`] for strategy styles.
///
/// Mirrors nesting Java `WriteFont` into `WriteCellStyle.writeFont`. Owned
/// `font_name` strings cannot become `&'static str`; pass name via
/// [`ExcelFontStyle::font_name`] when a static label is available. All other
/// common fields (size, color, bold, italic, underline, strikeout, charset,
/// type offset) are preserved.
#[must_use]
pub fn excel_font_style_from_write_font(font: &WriteFont) -> ExcelFontStyle {
    ExcelFontStyle {
        font_name: None,
        font_height_in_points: font.get_font_height_in_points(),
        italic: font.get_italic(),
        strikeout: font.get_strikeout(),
        color: font.get_color(),
        type_offset: font.get_type_offset(),
        underline: font.get_underline(),
        charset: font.get_charset(),
        bold: font.get_bold(),
    }
}
