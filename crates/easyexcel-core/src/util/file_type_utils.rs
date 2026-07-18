//! Mirrors Java com.alibaba.excel.util.FileTypeUtils.

#![allow(dead_code)]

/// Mirrors `com.alibaba.excel.util.FileTypeUtils#getImageTypeFormat`.
///
/// Returns the canonical lowercase file extension (without leading dot)
/// for a given image type name. Java normalises to `jpg`, `png`, `gif`,
/// `bmp`; an unknown name is returned unchanged.
#[must_use]
pub fn get_image_type_format(image_type: &str) -> String {
    let lower = image_type.to_ascii_lowercase();
    match lower.as_str() {
        "jpeg" | "jpg" => "jpg".to_owned(),
        "png" => "png".to_owned(),
        "gif" => "gif".to_owned(),
        "bmp" => "bmp".to_owned(),
        other => other.to_owned(),
    }
}

/// Mirrors `com.alibaba.excel.util.FileTypeUtils#getImageType`.
///
/// Sniffs the image type from the magic bytes of a file header.
#[must_use]
pub fn get_image_type(image_header: &[u8]) -> Option<&'static str> {
    if image_header.len() >= 3 && image_header[0..3] == [0xFF, 0xD8, 0xFF] {
        return Some("jpg");
    }
    if image_header.len() >= 4 && image_header[0..4] == [0x89, 0x50, 0x4E, 0x47] {
        return Some("png");
    }
    if image_header.len() >= 3 && image_header[0..3] == [0x47, 0x49, 0x46] {
        return Some("gif");
    }
    if image_header.len() >= 2 && image_header[0..2] == [0x42, 0x4D] {
        return Some("bmp");
    }
    None
}
