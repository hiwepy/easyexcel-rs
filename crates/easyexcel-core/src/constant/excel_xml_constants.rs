//! Mirrors Java `com.alibaba.excel.constant.ExcelXmlConstants`.

/// `dimension` tag.
pub const DIMENSION_TAG: &str = "dimension";
/// `row` tag.
pub const ROW_TAG: &str = "row";
/// `f` (formula) tag.
pub const CELL_FORMULA_TAG: &str = "f";
/// `v` (value) tag.
pub const CELL_VALUE_TAG: &str = "v";
/// `t` (inline string value) tag.
pub const CELL_INLINE_STRING_VALUE_TAG: &str = "t";
/// `c` (cell) tag.
pub const CELL_TAG: &str = "c";
/// `mergeCell` tag.
pub const MERGE_CELL_TAG: &str = "mergeCell";
/// `hyperlink` tag.
pub const HYPERLINK_TAG: &str = "hyperlink";

/// `s` attribute.
pub const ATTRIBUTE_S: &str = "s";
/// `ref` attribute.
pub const ATTRIBUTE_REF: &str = "ref";
/// `r` attribute.
pub const ATTRIBUTE_R: &str = "r";
/// `t` attribute.
pub const ATTRIBUTE_T: &str = "t";
/// `location` attribute.
pub const ATTRIBUTE_LOCATION: &str = "location";
/// `r:id` attribute.
pub const ATTRIBUTE_RID: &str = "r:id";

/// Cell range split character.
pub const CELL_RANGE_SPLIT: &str = ":";

// SharedStrings tags
/// `t` tag in shared strings.
pub const SHAREDSTRINGS_T_TAG: &str = "t";
/// `si` tag in shared strings.
pub const SHAREDSTRINGS_SI_TAG: &str = "si";
/// `rPh` tag (phonetic) in shared strings.
pub const SHAREDSTRINGS_RPH_TAG: &str = "rPh";
