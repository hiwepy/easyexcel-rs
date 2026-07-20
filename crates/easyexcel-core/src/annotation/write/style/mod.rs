//! Mirrors Java `com.alibaba.excel.annotation.write.style.*`.
//!
//! Java applies these annotations at runtime via `ExcelWriteHeadProperty` /
//! `AbstractWriteHolder`. In Rust, `#[derive(ExcelRow)]` with nested
//! `#[excel(...)]` attributes emits the same metadata on
//! [`crate::ExcelColumn`] and [`crate::ExcelWriteMetadata`].
//!
//! | Java annotation | Rust attribute (via `#[excel(...)]`) | Scope | Writer |
//! |---|---|---|---|
//! | `@ColumnWidth` | `column_width = N` | type / field | wired |
//! | `@HeadRowHeight` | `head_row_height = N` | type | wired |
//! | `@ContentRowHeight` | `content_row_height = N` | type | wired |
//! | `@HeadStyle` | `head_style(...)` | type / field | wired |
//! | `@ContentStyle` | `content_style(...)` | type / field | wired |
//! | `@HeadFontStyle` | `head_font_style(...)` | type / field | wired |
//! | `@ContentFontStyle` | `content_font_style(...)` | type / field | wired |
//! | `@ContentLoopMerge` | `content_loop_merge(...)` | field | wired |
//! | `@OnceAbsoluteMerge` | `once_absolute_merge(...)` | type | wired |
//!
//! Strategy types such as `HorizontalCellStyleStrategy` /
//! `SimpleColumnWidthStyleStrategy` are applied when registered as write
//! handlers (`WriteHandler::style_*` accessors). Annotations /
//! `WriteOptions` remain available as an alternative path. See
//! `easyexcel-writer::style` module docs.

pub mod column_width;
pub mod content_font_style;
pub mod content_loop_merge;
pub mod content_row_height;
pub mod content_style;
pub mod head_font_style;
pub mod head_row_height;
pub mod head_style;
pub mod once_absolute_merge;

pub use column_width::ColumnWidth;
pub use content_font_style::ContentFontStyle;
pub use content_loop_merge::ContentLoopMerge;
pub use content_row_height::ContentRowHeight;
pub use content_style::ContentStyle;
pub use head_font_style::HeadFontStyle;
pub use head_row_height::HeadRowHeight;
pub use head_style::HeadStyle;
pub use once_absolute_merge::OnceAbsoluteMerge;
