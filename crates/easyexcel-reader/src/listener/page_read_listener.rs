//! Mirrors Java `com.alibaba.excel.read.listener.PageReadListener`.

use std::collections::VecDeque;

use easyexcel_core::{AnalysisContext, ReadListener};

/// Mirrors Java `PageReadListener<T> implements ReadListener<T>`.
///
/// Java batches rows in a list and invokes a `Consumer<List<T>>` when the
/// batch is full or on `doAfterAllAnalysed`. Rust mirrors with a
/// `VecDeque<T>` and an injected callback.
pub struct PageReadListener<T> {
    batch_size: usize,
    rows: VecDeque<T>,
    callback: Box<dyn FnMut(Vec<T>) + Send>,
}

impl<T> PageReadListener<T> {
    /// Mirrors Java `PageReadListener(Consumer<List<T>>)`.
    pub fn new(callback: impl FnMut(Vec<T>) + Send + 'static) -> Self {
        Self {
            batch_size: 100,
            rows: VecDeque::new(),
            callback: Box::new(callback),
        }
    }

    /// Mirrors Java `PageReadListener(Consumer<List<T>>, int)`.
    pub fn with_batch_size(mut self, batch_size: usize) -> Self {
        self.batch_size = batch_size.max(1);
        self
    }

    /// Mirrors the static `BATCH_COUNT` constant.
    pub const BATCH_COUNT: usize = 100;

    /// Flushes any remaining rows. (Java `doAfterAllAnalysed` step)
    pub fn flush(&mut self) {
        if !self.rows.is_empty() {
            let drained: Vec<T> = self.rows.drain(..).collect();
            (self.callback)(drained);
        }
    }
}

impl<T: Send + 'static> ReadListener<T> for PageReadListener<T> {
    fn invoke(&mut self, data: T, _context: &AnalysisContext) -> easyexcel_core::Result<()> {
        self.rows.push_back(data);
        if self.rows.len() >= self.batch_size {
            self.flush();
        }
        Ok(())
    }

    fn do_after_all_analysed(&mut self, _context: &AnalysisContext) -> easyexcel_core::Result<()> {
        self.flush();
        Ok(())
    }
}