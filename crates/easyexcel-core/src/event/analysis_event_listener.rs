//! Mirrors Java `com.alibaba.excel.event.AnalysisEventListener`.

use std::collections::HashMap;

use crate::analysis_context::AnalysisContext;
use crate::cell_value::CellValue;
use crate::read_cell_data::ReadCellData;
use crate::read_listener::ReadListener;

/// Receives the return of each piece of data parsed.
///
/// Rust port of Java `AnalysisEventListener<T> implements ReadListener<T>`.
/// The Java side overrides `invokeHead` to convert `ReadCellData<?>` map
/// into a `String` map and delegate to `invokeHeadMap`. Rust mirrors
/// both methods.
pub trait AnalysisEventListener<T>: ReadListener<T> {
    /// Called for each header row after conversion to a `String` map.
    /// (Java `invokeHeadMap(Map<Integer, String>, AnalysisContext)`)
    fn invoke_head_map(
        &mut self,
        _head_map: &HashMap<usize, String>,
        _context: &AnalysisContext,
    ) {
    }
}

/// Adapter that converts `Map<Integer, ReadCellData<?>>` to
/// `Map<Integer, String>` before calling `invoke_head_map`.
/// (Java `AnalysisEventListener.invokeHead`)
pub fn convert_head_map(
    head_map: &HashMap<usize, ReadCellData>,
    _context: &AnalysisContext,
) -> HashMap<usize, String> {
    head_map
        .iter()
        .map(|(k, v)| (*k, v.display_value().to_owned()))
        .collect()
}

// Suppress unused import for CellValue.
#[allow(dead_code)]
fn _import_marker(_: CellValue) {}
