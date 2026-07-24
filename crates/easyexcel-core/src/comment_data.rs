//! Mirrors Java `com.alibaba.excel.metadata.data.CommentData`.

use crate::client_anchor_data::ClientAnchorData;
use crate::rich_text_string_data::RichTextStringData;

/// Cell comment metadata matching Java `CommentData extends ClientAnchorData`.
///
/// Rust uses composition for the anchor (same pattern as [`crate::ImageData`])
/// so `ClientAnchorData` stays `Copy`/`Default` without inheritance bookkeeping.
#[derive(Debug, Clone, Default, PartialEq)]
pub struct CommentData {
    author: Option<String>,
    rich_text_string_data: Option<RichTextStringData>,
    anchor: ClientAnchorData,
}

impl CommentData {
    /// Creates an empty comment. (Java default constructor)
    #[must_use]
    pub const fn new() -> Self {
        Self {
            author: None,
            rich_text_string_data: None,
            anchor: ClientAnchorData::new(),
        }
    }

    /// Sets the original comment author. (Java `setAuthor(String)`)
    #[must_use]
    pub fn author(mut self, author: impl Into<String>) -> Self {
        self.author = Some(author.into());
        self
    }

    /// Sets the rich-text body. (Java `setRichTextStringData(RichTextStringData)`)
    #[must_use]
    pub fn rich_text_string_data(mut self, value: RichTextStringData) -> Self {
        self.rich_text_string_data = Some(value);
        self
    }

    /// Sets plain-text body convenience (wraps [`RichTextStringData::new`]).
    #[must_use]
    pub fn text(mut self, text: impl Into<String>) -> Self {
        self.rich_text_string_data = Some(RichTextStringData::new(text));
        self
    }

    /// Sets the client anchor. (Java inherited `ClientAnchorData` fields)
    #[must_use]
    pub const fn anchor(mut self, value: ClientAnchorData) -> Self {
        self.anchor = value;
        self
    }

    /// Returns the author. (Java `getAuthor()`)
    #[must_use]
    pub fn get_author(&self) -> Option<&str> {
        self.author.as_deref()
    }

    /// Returns the rich-text body. (Java `getRichTextStringData()`)
    #[must_use]
    pub const fn get_rich_text_string_data(&self) -> Option<&RichTextStringData> {
        self.rich_text_string_data.as_ref()
    }

    /// Returns the client anchor.
    #[must_use]
    pub const fn get_anchor(&self) -> ClientAnchorData {
        self.anchor
    }

    /// Returns plain note text for writer backends that only accept a string.
    #[must_use]
    pub fn note_text(&self) -> String {
        self.rich_text_string_data
            .as_ref()
            .map(|r| r.text_string().to_owned())
            .unwrap_or_default()
    }
}
