//! 填充包装。
//!
//! 对应 Java：`com.alibaba.excel.write.metadata.fill.FillWrapper`
//! 权威实现位于 `easyexcel_template::FillWrapper`。

// 注意：writer 不直接依赖 template，避免环依赖；
// 请从 `easyexcel_template::FillWrapper` 或门面使用。

/// 迁移说明：FillWrapper 实现见 `easyexcel-template`。
pub struct FillWrapperPathDocs;
