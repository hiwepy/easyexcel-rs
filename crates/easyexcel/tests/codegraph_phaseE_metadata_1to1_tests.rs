//! Phase E — codegraph-driven 1:1 method-level tests for the metadata
//! classes expanded in Phase E.1..E.4.
//!
//! Java references:
//! - com.alibaba.excel.read.builder.ExcelReaderSheetBuilder
//! - com.alibaba.excel.read.metadata.ReadSheet
//! - com.alibaba.excel.read.metadata.ReadWorkbook
//! - com.alibaba.excel.write.metadata.WriteSheet
//! - com.alibaba.excel.write.metadata.WriteTable
//! - com.alibaba.excel.write.metadata.WriteWorkbook
//!
//! Naming: `mod <java_class_snake>` + `fn <java_method_snake>`.

// ---------------------------------------------------------------------------
// ExcelReaderSheetBuilder
// ---------------------------------------------------------------------------

mod excel_reader_sheet_builder_test {
    //! Mirrors ExcelReaderSheetBuilder constructors + setters
    use easyexcel_reader::builder::excel_reader_sheet_builder::ExcelReaderSheetBuilder;

    /// Java: `ExcelReaderSheetBuilder()` default ctor.
    #[test]
    fn t01_default_ctor07() {
        let b = ExcelReaderSheetBuilder::new();
        assert!(b.sheet_no.is_none());
        assert!(b.sheet_name.is_none());
    }

    /// Java: `sheetNo(Integer)` setter.
    #[test]
    fn t02_sheet_no_setter07() {
        let b = ExcelReaderSheetBuilder::new().sheet_no(3);
        assert_eq!(b.sheet_no, Some(3));
    }

    /// Java: `sheetName(String)` setter.
    #[test]
    fn t03_sheet_name_setter07() {
        let b = ExcelReaderSheetBuilder::new().sheet_name("Sheet2");
        assert_eq!(b.sheet_name.as_deref(), Some("Sheet2"));
    }

    /// Java: `headRowNumber(Integer)` inherited setter.
    #[test]
    fn t04_head_row_number_setter07() {
        let b = ExcelReaderSheetBuilder::new().head_row_number(2);
        assert_eq!(b.head_row_number, Some(2));
    }

    /// Java: `useScientificFormat(Boolean)` inherited setter.
    #[test]
    fn t05_use_scientific_format_setter07() {
        let b = ExcelReaderSheetBuilder::new().use_scientific_format(true);
        assert_eq!(b.use_scientific_format, Some(true));
    }

    /// Java: `build()` returns a ReadSheet with sheet_no set.
    #[test]
    fn t06_build_returns_read_sheet07() {
        let b = ExcelReaderSheetBuilder::new()
            .sheet_no(0)
            .sheet_name("Sheet1");
        let sheet = b.build();
        assert_eq!(sheet.sheet_no(), 0);
        assert_eq!(sheet.sheet_name(), "Sheet1");
    }

    /// Java: `parameter()` returns same shape as `build()`.
    #[test]
    fn t07_parameter_returns_read_sheet07() {
        let b = ExcelReaderSheetBuilder::new().sheet_no(2);
        assert_eq!(b.parameter().sheet_no(), 2);
    }

    /// Java: name-only sheet selection leaves `sheetNo` null and carries
    /// inherited read parameters into the built `ReadSheet`.
    #[test]
    fn t08_name_only_and_basic_parameters07() {
        let sheet = ExcelReaderSheetBuilder::new()
            .sheet_name("Sheet2")
            .head_row_number(2)
            .use_scientific_format(true)
            .build();
        assert!(!sheet.has_sheet_no());
        assert_eq!(sheet.sheet_name(), "Sheet2");
        assert_eq!(sheet.head_row_number(), Some(2));
        assert_eq!(sheet.use_scientific_format(), Some(true));
    }
}

// ---------------------------------------------------------------------------
// ReadSheet
// ---------------------------------------------------------------------------

mod read_sheet_test {
    //! Mirrors ReadSheet ctors + getters/setters + copyBasicParameter
    use easyexcel_reader::context::read_sheet::ReadSheet;

    /// Java: `ReadSheet()` no-arg ctor.
    #[test]
    fn t01_default_ctor07() {
        let s = ReadSheet::default_construction();
        assert_eq!(s.sheet_no(), 0);
        assert_eq!(s.sheet_name(), "");
    }

    /// Java: `ReadSheet(Integer sheetNo)`.
    #[test]
    fn t02_ctor_sheet_no_only07() {
        let s = ReadSheet::new(1);
        assert_eq!(s.sheet_no(), 1);
        assert_eq!(s.sheet_name(), "");
    }

    /// Java: `ReadSheet(Integer sheetNo, String sheetName)`.
    #[test]
    fn t03_ctor_sheet_no_and_name07() {
        let s = ReadSheet::with_name(0, "Data");
        assert_eq!(s.sheet_no(), 0);
        assert_eq!(s.sheet_name(), "Data");
    }

    /// Java: `getSheetNo()`.
    #[test]
    fn t04_get_sheet_no07() {
        let s = ReadSheet::new(7);
        assert_eq!(s.sheet_no(), 7);
    }

    /// Java: `setSheetNo(Integer)`.
    #[test]
    fn t05_set_sheet_no07() {
        let mut s = ReadSheet::default_construction();
        s.set_sheet_no(5);
        assert_eq!(s.sheet_no(), 5);
    }

    /// Java: `getSheetName()`.
    #[test]
    fn t06_get_sheet_name07() {
        let s = ReadSheet::with_name(0, "X");
        assert_eq!(s.sheet_name(), "X");
    }

    /// Java: `setSheetName(String)`.
    #[test]
    fn t07_set_sheet_name07() {
        let mut s = ReadSheet::default_construction();
        s.set_sheet_name("Y");
        assert_eq!(s.sheet_name(), "Y");
    }

    /// Java: `copyBasicParameter(ReadSheet other)`.
    #[test]
    fn t08_copy_basic_parameter07() {
        let mut a = ReadSheet::with_name(1, "A");
        let mut b = ReadSheet::with_name(2, "B");
        b.set_head_row_number(2).set_use_scientific_format(true);
        a.copy_basic_parameter(&b);
        assert_eq!(a.sheet_no(), 1);
        assert_eq!(a.sheet_name(), "A");
        assert_eq!(a.head_row_number(), Some(2));
        assert_eq!(a.use_scientific_format(), Some(true));
    }

    /// Java: `toString()` format.
    #[test]
    fn t09_to_string07() {
        let s = ReadSheet::with_name(1, "S1");
        let s_str = format!("{s}");
        assert!(s_str.contains("sheetNo=1"));
        assert!(s_str.contains("sheetName='S1'"));
    }
}

// ---------------------------------------------------------------------------
// ReadTable
// ---------------------------------------------------------------------------

mod read_table_test {
    //! Mirrors ReadTable (4 members)
    use easyexcel_reader::metadata::read_table::ReadTable;

    /// Java: `ReadTable()` no-arg ctor.
    #[test]
    fn t01_default_ctor07() {
        let t = ReadTable::new();
        assert_eq!(t.table_no(), 0);
    }

    /// Java: `ReadTable(Integer tableNo)`.
    #[test]
    fn t02_ctor_table_no07() {
        let t = ReadTable::with_table_no(3);
        assert_eq!(t.table_no(), 3);
    }

    /// Java: `getTableNo()`.
    #[test]
    fn t03_get_table_no07() {
        let t = ReadTable::with_table_no(7);
        assert_eq!(t.table_no(), 7);
    }

    /// Java: `setTableNo(Integer)`.
    #[test]
    fn t04_set_table_no07() {
        let mut t = ReadTable::new();
        t.set_table_no(5);
        assert_eq!(t.table_no(), 5);
    }
}

// ---------------------------------------------------------------------------
// ReadWorkbook
// ---------------------------------------------------------------------------

mod read_workbook_test {
    //! Mirrors ReadWorkbook
    use easyexcel_reader::metadata::read_workbook::ReadWorkbook;

    /// Java: `ReadWorkbook()` no-arg ctor.
    #[test]
    fn t01_default_ctor07() {
        let w = ReadWorkbook::new();
        // Java default: ReadBasicParameter's `ignoreEmptyRow` defaults
        // to true. Rust's ReadOptions also defaults to true, so the
        // mirror should match.
        assert!(w.ignore_empty_row());
    }

    /// Java: `getIgnoreEmptyRow()`.
    #[test]
    fn t02_get_ignore_empty_row07() {
        let w = ReadWorkbook::new();
        assert!(w.ignore_empty_row());
    }

    /// Java: `setIgnoreEmptyRow(Boolean)`.
    #[test]
    fn t03_set_ignore_empty_row07() {
        let mut w = ReadWorkbook::new();
        w.set_ignore_empty_row(false);
        assert!(!w.ignore_empty_row());
        w.set_ignore_empty_row(true);
        assert!(w.ignore_empty_row());
    }

    /// Java: `setCustomObject(Object)`.
    #[test]
    fn t04_set_custom_object07() {
        let mut w = ReadWorkbook::new();
        let co = easyexcel_core::CustomReadObject::new("payload");
        w.set_custom_object(co);
        assert!(w.custom_object().is_some());
    }

    /// Java: `setCharset(Charset)`.
    #[test]
    fn t05_set_charset07() {
        let mut w = ReadWorkbook::new();
        let cs = easyexcel_core::CsvCharset::default();
        w.set_charset(cs.clone());
        assert_eq!(w.charset(), &cs);
    }

    /// Java: `setPassword(String)`.
    #[test]
    fn t06_set_password07() {
        let mut w = ReadWorkbook::new();
        w.set_password("secret");
        assert_eq!(w.password(), Some("secret"));
    }

    /// Java: `setHeadRowNumber(Integer)`.
    #[test]
    fn t07_set_head_row_number07() {
        let mut w = ReadWorkbook::new();
        w.set_head_row_number(3);
        assert_eq!(w.head_row_number(), 3);
    }

    /// Java: `getReadCache()`.
    #[test]
    fn t08_get_read_cache07() {
        let w = ReadWorkbook::new();
        let _ = w.read_cache();
    }

    /// Java: `setReadCache(ReadCache)`.
    #[test]
    fn t09_set_read_cache07() {
        let mut w = ReadWorkbook::new();
        w.set_read_cache(easyexcel_reader::ReadCacheMode::Auto);
    }
}

// ---------------------------------------------------------------------------
// WriteSheet
// ---------------------------------------------------------------------------

mod write_sheet_test {
    //! Mirrors WriteSheet setters
    use easyexcel_writer::MirroredWriteSheet;

    /// Java: `getSheetNo()` / `setSheetNo(Integer)`.
    #[test]
    fn t01_sheet_no07() {
        let mut s = MirroredWriteSheet::new();
        assert_eq!(s.sheet_no(), 0);
        s.set_sheet_no(5);
        assert_eq!(s.sheet_no(), 5);
    }

    /// Java: `getSheetName()` / `setSheetName(String)`.
    #[test]
    fn t02_sheet_name07() {
        let mut s = MirroredWriteSheet::new();
        s.set_sheet_name("X");
        assert_eq!(s.sheet_name(), "X");
    }
}

// ---------------------------------------------------------------------------
// WriteTable
// ---------------------------------------------------------------------------

mod write_table_test {
    //! Mirrors WriteTable setter
    use easyexcel_writer::MirroredWriteTable;

    /// Java: `getTableNo()` / `setTableNo(Integer)`.
    #[test]
    fn t01_table_no07() {
        let mut t = MirroredWriteTable::new();
        assert_eq!(t.table_no(), 0);
        t.set_table_no(3);
        assert_eq!(t.table_no(), 3);
    }
}

// ---------------------------------------------------------------------------
// WriteWorkbook
// ---------------------------------------------------------------------------

mod write_workbook_test {
    //! Mirrors WriteWorkbook
    use easyexcel_writer::MirroredWriteWorkbook;

    /// Java: `getExcelType()` / `setExcelType(ExcelTypeEnum)`.
    #[test]
    fn t01_excel_type07() {
        let mut w = MirroredWriteWorkbook::new();
        let initial = w.excel_type();
        w.set_excel_type(easyexcel_core::support::ExcelTypeEnum::Xls);
        assert_eq!(w.excel_type(), easyexcel_core::support::ExcelTypeEnum::Xls);
        let _ = initial;
    }

    /// Java: `getWithBom()` / `setWithBom(Boolean)`.
    #[test]
    fn t02_with_bom07() {
        let mut w = MirroredWriteWorkbook::new();
        w.set_with_bom(true);
        assert!(w.with_bom());
    }

    /// Java: `getPassword()` / `setPassword(String)`.
    #[test]
    fn t03_password07() {
        let mut w = MirroredWriteWorkbook::new();
        w.set_password("p");
        assert_eq!(w.password(), Some("p"));
    }

    /// Java: `getInMemory()` / `setInMemory(boolean)`.
    #[test]
    fn t04_in_memory07() {
        let mut w = MirroredWriteWorkbook::new();
        w.set_in_memory(true);
        assert!(w.in_memory());
    }

    /// Java: `getWriteExcelOnException()` / `setWriteExcelOnException(boolean)`.
    #[test]
    fn t05_write_excel_on_exception07() {
        let mut w = MirroredWriteWorkbook::new();
        w.set_write_excel_on_exception(true);
        assert!(w.write_excel_on_exception());
    }

    /// Java: `getAutoCloseStream()` / `setAutoCloseStream(boolean)`.
    #[test]
    fn t06_auto_close_stream07() {
        let mut w = MirroredWriteWorkbook::new();
        w.set_auto_close_stream(false);
        assert!(!w.auto_close_stream());
    }
}

// ---------------------------------------------------------------------------
// ExcelReaderTableBuilder (legacy compatibility; absent from Java 4.0.3)
// ---------------------------------------------------------------------------

mod excel_reader_table_builder_test {
    //! Preserves the pre-4.x `ExcelReaderTableBuilder` compatibility surface.
    use easyexcel_reader::builder::excel_reader_table_builder::ExcelReaderTableBuilder;

    /// Legacy Java: `ExcelReaderTableBuilder()` default ctor.
    #[test]
    fn t01_default_ctor07() {
        let b = ExcelReaderTableBuilder::new();
        assert!(b.table_no.is_none());
    }

    /// Legacy Java: `tableNo(Integer)` setter.
    #[test]
    fn t02_table_no_setter07() {
        let b = ExcelReaderTableBuilder::new().table_no(2);
        assert_eq!(b.table_no, Some(2));
    }

    /// Legacy Java: `build()` returns a ReadTable.
    #[test]
    fn t03_build_returns_read_table07() {
        let b = ExcelReaderTableBuilder::new().table_no(3);
        let t = b.build();
        assert_eq!(t.table_no(), 3);
    }

    /// Legacy Java: `parameter()` returns same as `build()`.
    #[test]
    fn t04_parameter_returns_read_table07() {
        let b = ExcelReaderTableBuilder::new().table_no(4);
        assert_eq!(b.parameter().table_no(), 4);
    }

    /// Legacy Java: `headRowNumber(Integer)` inherited setter.
    #[test]
    fn t05_head_row_number_setter07() {
        let b = ExcelReaderTableBuilder::new().head_row_number(1);
        assert_eq!(b.head_row_number, Some(1));
    }

    /// Legacy Java: `useScientificFormat(Boolean)` inherited setter.
    #[test]
    fn t06_use_scientific_format_setter07() {
        let b = ExcelReaderTableBuilder::new().use_scientific_format(true);
        assert_eq!(b.use_scientific_format, Some(true));
    }
}
