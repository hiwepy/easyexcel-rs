//! Mirrors Java `com.alibaba.excel.write.metadata.style.WriteFont`.

use easyexcel_core::WriteFont;

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