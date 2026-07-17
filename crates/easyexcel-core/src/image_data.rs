//! Mirrors Java `com.alibaba.excel.metadata.data.ImageData`.

use crate::client_anchor_data::ClientAnchorData;
use crate::image_type::ImageType;

/// One Java-compatible image and its client anchor.
///
/// Java `ImageData extends ClientAnchorData`; Rust uses composition for the
/// same reason as `ClientAnchorData`.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct ImageData {
    image: Vec<u8>,
    image_type: Option<ImageType>,
    anchor: ClientAnchorData,
}

impl ImageData {
    /// Creates image data from encoded bytes. (Java `ImageData(byte[])`)
    #[must_use]
    pub fn new(image: impl Into<Vec<u8>>) -> Self {
        Self {
            image: image.into(),
            image_type: None,
            anchor: ClientAnchorData::new(),
        }
    }

    /// Sets optional Java image-type metadata.
    #[must_use]
    pub const fn image_type(mut self, value: ImageType) -> Self {
        self.image_type = Some(value);
        self
    }

    /// Sets the client anchor.
    #[must_use]
    pub const fn anchor(mut self, value: ClientAnchorData) -> Self {
        self.anchor = value;
        self
    }

    /// Returns the encoded image bytes. (Java `getImage()`)
    #[must_use]
    pub fn image(&self) -> &[u8] {
        &self.image
    }

    /// Returns the optional image-type metadata. (Java `getImageType()`)
    #[must_use]
    pub const fn get_image_type(&self) -> Option<ImageType> {
        self.image_type
    }

    /// Returns the client anchor. (Java `getAnchor()`)
    #[must_use]
    pub const fn get_anchor(&self) -> ClientAnchorData {
        self.anchor
    }
}
