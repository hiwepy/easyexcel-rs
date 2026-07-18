//! Mirrors Java `com.alibaba.excel.write.executor.ExcelWriteExecutor` (interface).

/// Mirrors Java `ExcelWriteExecutor` (empty marker interface).
///
/// Java uses this interface to expose the abstract `ExcelWriteExecutor` to
/// other modules. The Rust port keeps the marker trait for 1:1 API
/// parity even though no methods need overriding.
pub trait ExcelWriteExecutor {}