//! Mirrors Java `com.alibaba.excel.read.listener.PageReadListener<T>`.

use crate::analysis_context::{AnalysisContext, Result};
use crate::read_listener::ReadListener;

/// A listener that buffers rows and invokes a callback page by page.
///
/// Mirrors Java `PageReadListener<T>(Consumer<List<T>> consumer, int
/// batchCount)` with `BATCH_COUNT = 100`.
pub struct PageReadListener<T> {
    batch_size: usize,
    batch_index: usize,
    rows: Vec<T>,
    callback: Box<PageCallback<T>>,
}

/// Callback signature for [`PageReadListener`].
type PageCallback<T> = dyn FnMut(Vec<T>, &AnalysisContext) -> Result<()>;

impl<T> PageReadListener<T> {
    /// Creates a paged listener. A zero size is normalized to one row. (Java `PageReadListener(Consumer, int)`)
    #[must_use]
    pub fn new(
        batch_size: usize,
        callback: impl FnMut(Vec<T>, &AnalysisContext) -> Result<()> + 'static,
    ) -> Self {
        let batch_size = batch_size.max(1);
        Self {
            batch_size,
            batch_index: 0,
            rows: Vec::with_capacity(batch_size),
            callback: Box::new(callback),
        }
    }

    fn flush(&mut self, context: &AnalysisContext) -> Result<()> {
        if self.rows.is_empty() {
            return Ok(());
        }
        let rows = std::mem::replace(&mut self.rows, Vec::with_capacity(self.batch_size));
        let context = context.with_batch_index(self.batch_index);
        complete_page(&mut self.batch_index, (self.callback)(rows, &context))
    }
}

impl<T> ReadListener<T> for PageReadListener<T> {
    fn invoke(&mut self, data: T, context: &AnalysisContext) -> Result<()> {
        self.rows.push(data);
        if self.rows.len() >= self.batch_size {
            return self.flush(context);
        }
        Ok(())
    }

    fn do_after_all_analysed(&mut self, context: &AnalysisContext) -> Result<()> {
        self.flush(context)
    }
}

fn complete_page(batch_index: &mut usize, result: Result<()>) -> Result<()> {
    result.map(|()| {
        *batch_index += 1;
    })
}
