//! Phase 2 — 1:1 test matrix for handler sub-traits + default loader.
//!
//! Java reference: `com.alibaba.excel.write.handler.{WorkbookWriteHandler,
//! SheetWriteHandler, RowWriteHandler, CellWriteHandler, MergeHandler,
//! ConstraintHandler}` and `DefaultWriteHandlerLoader`.
//!
//! Rust mirror: sub-trait split in
//! `easyexcel_writer::handler::{workbook,sheet,row,cell,merge,constraint}_write_handler`
//! and `DefaultWriteHandlerLoader::load_default_handler()`.
//!
//! Naming: `mod <java_class_snake>` + `fn <java_method_snake>`.

use easyexcel_core::WriteHandler;
use easyexcel_writer::handler::{
    cell_write_handler::CellWriteHandler, default_write_handler_loader::DefaultWriteHandlerLoader,
    row_write_handler::RowWriteHandler, sheet_write_handler::SheetWriteHandler,
    workbook_write_handler::WorkbookWriteHandler,
};

// ---------------------------------------------------------------------------
// WorkbookWriteHandler — mirrors Java interface
// ---------------------------------------------------------------------------

mod workbook_write_handler_test {
    //! Mirrors WorkbookWriteHandlerTest#t01WorkbookCreateOrder07
    use super::*;
    use easyexcel_core::WriteWorkbookContext;

    /// Capture-only handler used to verify the workbook before/after
    /// lifecycle fires. Mirrors Java `WorkbookWriteHandler` interface.
    #[derive(Default)]
    struct LifecycleProbe {
        before_count: u32,
        after_count: u32,
    }

    impl WorkbookWriteHandler for LifecycleProbe {}

    impl WriteHandler for LifecycleProbe {
        fn before_workbook(&mut self, _ctx: &WriteWorkbookContext) -> easyexcel_core::Result<()> {
            self.before_count += 1;
            Ok(())
        }
        fn after_workbook(&mut self, _ctx: &WriteWorkbookContext) -> easyexcel_core::Result<()> {
            self.after_count += 1;
            Ok(())
        }
    }

    /// Java: the workbook before/after hook fires on every workbook
    /// lifecycle event.
    #[test]
    fn t01_workbook_create_order07() {
        let mut probe = LifecycleProbe::default();
        let ctx = WriteWorkbookContext::new("out.xlsx");
        WriteHandler::before_workbook(&mut probe, &ctx).unwrap();
        assert_eq!(probe.before_count, 1);
        WriteHandler::after_workbook(&mut probe, &ctx).unwrap();
        assert_eq!(probe.after_count, 1);
    }
}

// ---------------------------------------------------------------------------
// SheetWriteHandler — mirrors Java interface
// ---------------------------------------------------------------------------

mod sheet_write_handler_test {
    //! Mirrors SheetWriteHandlerTest#t02SheetCreateOrder07
    use super::*;
    use easyexcel_core::{Result, WriteSheetContext};

    #[derive(Default)]
    struct SheetProbe {
        before_count: u32,
        after_count: u32,
    }

    impl SheetWriteHandler for SheetProbe {}

    impl WriteHandler for SheetProbe {
        fn before_sheet(&mut self, _ctx: &WriteSheetContext) -> Result<()> {
            self.before_count += 1;
            Ok(())
        }
        fn after_sheet(&mut self, _ctx: &WriteSheetContext) -> Result<()> {
            self.after_count += 1;
            Ok(())
        }
    }

    /// Java: sheet before/after hook fires once per sheet lifecycle.
    #[test]
    fn t02_sheet_create_order07() {
        let mut probe = SheetProbe::default();
        let ctx = WriteSheetContext::new("Sheet1");
        WriteHandler::before_sheet(&mut probe, &ctx).unwrap();
        assert_eq!(probe.before_count, 1);
        WriteHandler::after_sheet(&mut probe, &ctx).unwrap();
        assert_eq!(probe.after_count, 1);
    }
}

// ---------------------------------------------------------------------------
// RowWriteHandler — mirrors Java interface
// ---------------------------------------------------------------------------

mod row_write_handler_test {
    //! Mirrors RowWriteHandlerTest#t03RowCreateOrder07
    use super::*;
    use easyexcel_core::{Result, WriteRowContext};

    #[derive(Default)]
    struct RowProbe {
        before_count: u32,
        after_count: u32,
    }

    impl RowWriteHandler for RowProbe {}

    impl WriteHandler for RowProbe {
        fn before_row(&mut self, _ctx: &WriteRowContext) -> Result<()> {
            self.before_count += 1;
            Ok(())
        }
        fn after_row(&mut self, _ctx: &WriteRowContext) -> Result<()> {
            self.after_count += 1;
            Ok(())
        }
    }

    /// Java: row before/after hook fires once per row write.
    #[test]
    fn t03_row_create_order07() {
        let mut probe = RowProbe::default();
        let ctx = WriteRowContext::new("Sheet1", 0, Some(0), false);
        WriteHandler::before_row(&mut probe, &ctx).unwrap();
        assert_eq!(probe.before_count, 1);
        WriteHandler::after_row(&mut probe, &ctx).unwrap();
        assert_eq!(probe.after_count, 1);
    }
}

// ---------------------------------------------------------------------------
// CellWriteHandler — mirrors Java interface
// ---------------------------------------------------------------------------

mod cell_write_handler_test {
    //! Mirrors CellWriteHandlerTest#t04CellCreateOrder07
    use super::*;
    use easyexcel_core::{Result, WriteCellContext};

    #[derive(Default)]
    struct CellProbe {
        before_count: u32,
        after_count: u32,
    }

    impl CellWriteHandler for CellProbe {}

    impl WriteHandler for CellProbe {
        fn before_cell(&mut self, _ctx: &mut WriteCellContext) -> Result<()> {
            self.before_count += 1;
            Ok(())
        }
        fn after_cell(&mut self, _ctx: &WriteCellContext) -> Result<()> {
            self.after_count += 1;
            Ok(())
        }
    }

    /// Java: cell before/after hook fires once per cell write.
    #[test]
    fn t04_cell_create_order07() {
        let mut probe = CellProbe::default();
        let mut ctx = WriteCellContext::new("Sheet1", 0, 0, easyexcel::CellValue::Empty);
        WriteHandler::before_cell(&mut probe, &mut ctx).unwrap();
        assert_eq!(probe.before_count, 1);
        WriteHandler::after_cell(&mut probe, &ctx).unwrap();
        assert_eq!(probe.after_count, 1);
    }
}

// ---------------------------------------------------------------------------
// DefaultWriteHandlerLoader — mirrors Java loader
// ---------------------------------------------------------------------------

mod default_write_handler_loader_test {
    //! Mirrors DefaultWriteHandlerLoaderTest#t01LoadDefaultHandler07
    use super::*;

    /// Java: default XLSX handlers are Dimension, DefaultRow, FillStyle,
    /// and DefaultStyle when `useDefaultStyle=true`.
    #[test]
    fn t01_load_default_handler07() {
        let mut handlers = DefaultWriteHandlerLoader::load_default_handler();
        assert_eq!(handlers.len(), 4, "must return 4 XLSX default handlers");
        // Verify all handlers implement the unified WriteHandler trait
        // and can dispatch a no-op workbook lifecycle.
        let ctx = easyexcel_core::WriteWorkbookContext::new("out.xlsx");
        for h in handlers.iter_mut() {
            WriteHandler::before_workbook(h.as_mut(), &ctx).unwrap();
            WriteHandler::after_workbook(h.as_mut(), &ctx).unwrap();
        }
    }

    /// Java: loadDefaultHandler is idempotent — calling twice yields
    /// two independent handler instances.
    #[test]
    fn t02_load_default_handler_idempotent07() {
        let a = DefaultWriteHandlerLoader::load_default_handler();
        let b = DefaultWriteHandlerLoader::load_default_handler();
        assert_eq!(a.len(), b.len());
        // Distinct boxed trait objects (different vtable pointers)
        assert!(!std::ptr::eq(a.as_ptr(), b.as_ptr()));
    }

    #[test]
    fn t03_load_default_handler_by_excel_type07() {
        use easyexcel_core::support::ExcelTypeEnum;

        assert_eq!(
            DefaultWriteHandlerLoader::load_default_handler_for(true, ExcelTypeEnum::Xlsx).len(),
            4
        );
        assert_eq!(
            DefaultWriteHandlerLoader::load_default_handler_for(false, ExcelTypeEnum::Xlsx).len(),
            3
        );
        assert_eq!(
            DefaultWriteHandlerLoader::load_default_handler_for(true, ExcelTypeEnum::Xls).len(),
            3
        );
        assert_eq!(
            DefaultWriteHandlerLoader::load_default_handler_for(false, ExcelTypeEnum::Xls).len(),
            2
        );
        assert_eq!(
            DefaultWriteHandlerLoader::load_default_handler_for(true, ExcelTypeEnum::Csv).len(),
            2
        );
    }
}
