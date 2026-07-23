//! Excel 内部单元格数据模型。
//!
//! 对应 Java：`com.alibaba.excel.metadata.data.CellData`
//! 原文件：`easyexcel-core/.../metadata/data/CellData.java`
//!
//! 读路径既有 [`crate::ReadCellData`]，写路径既有 [`crate::WriteCellData`]；
//! 本结构保留 Java 泛型 `CellData<T>` 的字段面，便于 1:1 迁移与测试对照。

use bigdecimal::BigDecimal;

use crate::CellDataType;
use crate::FormulaData;

/// Excel 内部单元格数据，对齐 Java `CellData<T>`。
///
/// # Java 对应字段
/// - `type` → [`Self::cell_type`]
/// - `numberValue` → [`Self::number_value`]
/// - `stringValue` → [`Self::string_value`]
/// - `booleanValue` → [`Self::boolean_value`]
/// - `data` → [`Self::data`]
/// - `formulaData` → [`Self::formula_data`]
#[derive(Debug, Clone, PartialEq)]
pub struct CellData<T = ()> {
    /// 单元格类型。Java `type` / `getType()` / `setType`
    pub cell_type: Option<CellDataType>,
    /// 数值。Java `numberValue`
    pub number_value: Option<BigDecimal>,
    /// 字符串或错误文本。Java `stringValue`
    pub string_value: Option<String>,
    /// 布尔值。Java `booleanValue`
    pub boolean_value: Option<bool>,
    /// 转换后的业务数据。Java `data`
    pub data: Option<T>,
    /// 公式。Java `formulaData`
    pub formula_data: Option<FormulaData>,
    /// 行号。来自 Java `AbstractCell.rowIndex`
    pub row_index: Option<usize>,
    /// 列号。来自 Java `AbstractCell.columnIndex`
    pub column_index: Option<usize>,
}

impl<T> Default for CellData<T> {
    fn default() -> Self {
        Self {
            cell_type: None,
            number_value: None,
            string_value: None,
            boolean_value: None,
            data: None,
            formula_data: None,
            row_index: None,
            column_index: None,
        }
    }
}

impl<T> CellData<T> {
    /// 创建空单元格。对应 Java 默认构造。
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }

    /// 确保类型非空并按值收缩 EMPTY。对应 Java `checkEmpty()`。
    pub fn check_empty(&mut self) {
        if self.cell_type.is_none() {
            self.cell_type = Some(CellDataType::Empty);
        }
        match self.cell_type {
            Some(CellDataType::String | CellDataType::DirectString | CellDataType::Error) => {
                if self.string_value.as_ref().is_none_or(String::is_empty) {
                    self.cell_type = Some(CellDataType::Empty);
                }
            }
            Some(CellDataType::Number) => {
                if self.number_value.is_none() {
                    self.cell_type = Some(CellDataType::Empty);
                }
            }
            Some(CellDataType::Boolean) => {
                if self.boolean_value.is_none() {
                    self.cell_type = Some(CellDataType::Empty);
                }
            }
            _ => {}
        }
    }
}
