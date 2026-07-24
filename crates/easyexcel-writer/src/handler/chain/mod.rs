//! Mirrors Java `com.alibaba.excel.write.handler.chain.*`.

pub mod cell_handler_execution_chain;
pub mod row_handler_execution_chain;
pub mod sheet_handler_execution_chain;
pub mod workbook_handler_execution_chain;

#[cfg(test)]
mod tests {
    use std::sync::{Arc, Mutex};

    use easyexcel_core::{
        CellValue, Result, WriteCellContext, WriteHandler, WriteRowContext, WriteSheetContext,
        WriteWorkbookContext,
    };

    use super::cell_handler_execution_chain::CellHandlerExecutionChain;
    use super::row_handler_execution_chain::RowHandlerExecutionChain;
    use super::sheet_handler_execution_chain::SheetHandlerExecutionChain;
    use super::workbook_handler_execution_chain::WorkbookHandlerExecutionChain;

    struct Probe {
        name: &'static str,
        events: Arc<Mutex<Vec<String>>>,
    }

    impl Probe {
        fn boxed(name: &'static str, events: &Arc<Mutex<Vec<String>>>) -> Box<dyn WriteHandler> {
            Box::new(Self {
                name,
                events: Arc::clone(events),
            })
        }

        fn push(&self, event: &str) {
            self.events
                .lock()
                .expect("chain event mutex poisoned")
                .push(format!("{}:{event}", self.name));
        }
    }

    impl WriteHandler for Probe {
        fn before_workbook_create(&mut self, _context: &WriteWorkbookContext) -> Result<()> {
            self.push("before_workbook");
            Ok(())
        }

        fn after_workbook_create(&mut self, _context: &WriteWorkbookContext) -> Result<()> {
            self.push("after_workbook");
            Ok(())
        }

        fn after_workbook_dispose(&mut self, _context: &WriteWorkbookContext) -> Result<()> {
            self.push("dispose_workbook");
            Ok(())
        }

        fn before_sheet_create(&mut self, _context: &WriteSheetContext) -> Result<()> {
            self.push("before_sheet");
            Ok(())
        }

        fn after_sheet_create(&mut self, _context: &WriteSheetContext) -> Result<()> {
            self.push("after_sheet");
            Ok(())
        }

        fn before_row_create(&mut self, _context: &WriteRowContext) -> Result<()> {
            self.push("before_row");
            Ok(())
        }

        fn after_row_create(&mut self, _context: &WriteRowContext) -> Result<()> {
            self.push("after_row");
            Ok(())
        }

        fn after_row_dispose(&mut self, _context: &WriteRowContext) -> Result<()> {
            self.push("dispose_row");
            Ok(())
        }

        fn before_cell_create(&mut self, _context: &mut WriteCellContext) -> Result<()> {
            self.push("before_cell");
            Ok(())
        }

        fn after_cell_create(&mut self, _context: &WriteCellContext) -> Result<()> {
            self.push("after_cell");
            Ok(())
        }

        fn after_cell_data_converted(&mut self, _context: &WriteCellContext) -> Result<()> {
            self.push("converted_cell");
            Ok(())
        }

        fn after_cell_dispose(&mut self, _context: &WriteCellContext) -> Result<()> {
            self.push("dispose_cell");
            Ok(())
        }
    }

    #[test]
    fn all_java_chain_lifecycle_methods_forward_in_registration_order() -> Result<()> {
        let events = Arc::new(Mutex::new(Vec::new()));

        let mut workbook = WorkbookHandlerExecutionChain::with_handler(Probe::boxed("w1", &events));
        workbook.add_last(Probe::boxed("w2", &events));
        let workbook_context = WriteWorkbookContext::new("chain.xlsx");
        workbook.before_workbook_create(&workbook_context)?;
        workbook.after_workbook_create(&workbook_context)?;
        workbook.after_workbook_dispose(&workbook_context)?;

        let mut sheet = SheetHandlerExecutionChain::with_handler(Probe::boxed("s1", &events));
        sheet.add_last(Probe::boxed("s2", &events));
        let sheet_context = WriteSheetContext::new("Data");
        sheet.before_sheet_create(&sheet_context)?;
        sheet.after_sheet_create(&sheet_context)?;

        let mut row = RowHandlerExecutionChain::with_handler(Probe::boxed("r1", &events));
        row.add_last(Probe::boxed("r2", &events));
        let row_context = WriteRowContext::new("Data", 0, Some(0), false);
        row.before_row_create(&row_context)?;
        row.after_row_create(&row_context)?;
        row.after_row_dispose(&row_context)?;

        let mut cell = CellHandlerExecutionChain::with_handler(Probe::boxed("c1", &events));
        cell.add_last(Probe::boxed("c2", &events));
        let mut cell_context =
            WriteCellContext::new("Data", 0, 0, CellValue::String("value".to_owned()));
        cell.before_cell_create(&mut cell_context)?;
        cell.after_cell_create(&cell_context)?;
        cell.after_cell_data_converted(&cell_context)?;
        cell.after_cell_dispose(&cell_context)?;

        assert_eq!(
            *events.lock().expect("chain event mutex poisoned"),
            vec![
                "w1:before_workbook",
                "w2:before_workbook",
                "w1:after_workbook",
                "w2:after_workbook",
                "w1:dispose_workbook",
                "w2:dispose_workbook",
                "s1:before_sheet",
                "s2:before_sheet",
                "s1:after_sheet",
                "s2:after_sheet",
                "r1:before_row",
                "r2:before_row",
                "r1:after_row",
                "r2:after_row",
                "r1:dispose_row",
                "r2:dispose_row",
                "c1:before_cell",
                "c2:before_cell",
                "c1:after_cell",
                "c2:after_cell",
                "c1:converted_cell",
                "c2:converted_cell",
                "c1:dispose_cell",
                "c2:dispose_cell",
            ]
        );
        Ok(())
    }
}
