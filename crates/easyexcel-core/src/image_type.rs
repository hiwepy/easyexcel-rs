//! Mirrors Java `com.alibaba.excel.metadata.data.ImageData.ImageType`.

/// Java `ImageData.ImageType` equivalent metadata.
///
/// Java retains the POI numeric codes (2..=7). Rust drops them and maps to
/// `rust_xlsxwriter::Image` automatically; the enum is preserved for API
/// completeness.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ImageType {
    /// Extended Windows metafile.
    Emf,
    /// Windows metafile.
    Wmf,
    /// Macintosh PICT.
    Pict,
    /// JPEG.
    Jpeg,
    /// PNG.
    Png,
    /// Device-independent bitmap.
    Dib,
}
