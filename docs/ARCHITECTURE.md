# easyexcel-rs Architecture

> Rust 1:1 mirror of Alibaba EasyExcel 4.0.3. Covers reading, writing,
> template filling, converters, handlers, and encryption for XLSX/XLS/CSV.

## Crate Layout

```
easyexcel-rs/                        (workspace root)
├── crates/
│   ├── easyexcel/                   ← user-facing facade
│   ├── easyexcel-core/              ← traits, data models, errors
│   ├── easyexcel-derive/            ← proc-macro (`#[derive(ExcelRow)]`)
│   ├── easyexcel-reader/            ← XLSX/XLS/CSV reading
│   ├── easyexcel-template/          ← template fill (`{key}` placeholders)
│   └── easyexcel-writer/            ← XLSX/XLS/CSV writing + BIFF8 encoder
├── docs/
├── scripts/
└── tests/                           (integration tests in crate easyexcel)
```

## Data Flow

```
User Code
    │
    ▼
┌──────────────────┐
│   EasyExcel      │  ← facade (static factory: read / write / fill)
│   (easyexcel)    │
└──────┬───────────┘
       │
       ├──── read ────► ExcelReaderBuilder ──► ExcelReader
       │                     │                       │
       │                     ▼                       ▼
       │              ReadOptions            ExcelAnalyserImpl
       │                                            │
       │                     ┌──────────────────────┤
       │                     │ XLSX     │ XLS  │ CSV│
       │                     ▼          ▼       ▼   │
       │              XlsxSaxAnalyser  XlsSax  Csv  │
       │                                            │
       │    ┌─────── ReadListener ◄─────────────────┘
       │    │       (invoke / extra / on_exception)
       │    ▼
       │  User Row Type (T: ExcelRow)
       │
       ├──── write ───► ExcelWriterBuilder ──► ExcelWriter
       │                     │                      │
       │                     ▼                      │
       │              WriteOptions          ┌───────┤
       │                                    │XLSX│XLS│CSV
       │                                    ▼    ▼   ▼
       │                            rust_xlsxwriter biff8 csv
       │                                    │
       │    ┌─────── WriteHandler ◄─────────┘
       │    │      (before/after × workbook/sheet/row/cell)
       │    ▼
       │  Style / Merge / Width strategies
       │
       └──── fill ───► fill_xlsx_template / fill_xls_template_scalar
                           │
                           ▼
                    ExcelTemplateWriter (XLSX)
                    Biff8TemplatePackage (XLS)
                           │
                           ▼
                    Output XLSX / XLS / CSV
```

## Core Traits

| Trait | Location | Java Mirror |
|-------|----------|-------------|
| `ExcelRow` | `easyexcel-core` | `@ExcelProperty` + `ModelBuildEventListener` |
| `ReadListener<T>` | `easyexcel-core` | `com.alibaba.excel.read.listener.ReadListener` |
| `WriteHandler` | `easyexcel-core` | `Workbook/Sheet/Row/CellWriteHandler` |
| `Converter<T>` | `easyexcel-core` | `com.alibaba.excel.converters.Converter` |
| `IntoTemplateValue` | `easyexcel-template` | `FillWrapper` / `TemplateData` |
| `ReadCache` | `easyexcel-reader` | `com.alibaba.excel.cache.ReadCache` |

## File Format Support

| Feature | XLSX | XLS | CSV |
|---------|------|-----|-----|
| Read (typed rows) | ✅ | ✅ | ✅ |
| Read (dynamic / no-model) | ✅ | ✅ | ✅ |
| Read (event listener) | ✅ | ✅ | ✅ |
| Read (password-protected) | ✅ | ✅ RC4 | — |
| Write (typed rows) | ✅ | ✅ BIFF8 | ✅ |
| Write (with password) | ✅ Agile | ✅ RC4 | — |
| Write (constant memory / SXSSF) | ✅ | — | — |
| Template fill (`{key}`) | ✅ | ✅ LABEL | — |
| Template fill (list `{.}`) | ✅ | ✅ | — |
| Merge cells | ✅ | ✅ | — |
| Column width | ✅ | ✅ | — |
| Row height | ✅ | ✅ | — |
| Styles (font / fill / alignment) | ✅ | ✅ basic | — |
| Comments / Notes | ✅ read+write | ✅ read | — |
| Hyperlinks | ✅ read+write | ✅ read | — |
| Images | ✅ read+write | ✅ write | — |
| Formulas | ✅ read+write | — | — |
| Auto-filter | ✅ | — | — |

## Engine Dependencies

| Format | Read Engine | Write Engine |
|--------|------------|-------------|
| XLSX | Custom SAX parser (`quick-xml`) | `rust_xlsxwriter` |
| XLS | `calamine` + BIFF record handlers | Custom BIFF8 encoder |
| CSV | `csv` crate + `encoding_rs` | `csv` crate |
| Encryption (XLSX) | `office-crypto` | `ms-offcrypto-writer` (Agile) |
| Encryption (XLS) | Custom RC4 (`md-5` + `getrandom`) | Custom RC4 |
| ZIP (XLSX container) | `zip` crate | `zip` crate |
| OLE (XLS container) | `cfb` crate | `cfb` crate |

`calamine 0.36` remains the compatibility-oriented workbook backend, currently
used for legacy XLS reads. XLSX listener reads stay on the custom `quick-xml`
event pipeline because `worksheet_range` materializes a complete sheet and
`worksheet_range_at` selects a sheet by ordinal rather than reading a rectangular
chunk. ODS support is intentionally outside the Java EasyExcel compatibility
contract and can be added later as an opt-in extension without changing this
core pipeline.

## Derive Macro

`#[derive(ExcelRow)]` replaces Java's runtime annotation processing.
Supported attributes:

```rust
#[derive(ExcelRow)]
#[excel(ignore_unannotated)]           // @ExcelIgnoreUnannotated
#[excel(column_width = 20)]            // @ColumnWidth (type-level)
#[excel(head_row_height = 24)]         // @HeadRowHeight
#[excel(content_row_height = 16)]      // @ContentRowHeight
#[excel(head_style(...))]              // @HeadStyle
#[excel(content_style(...))]           // @ContentStyle
#[excel(head_font_style(...))]         // @HeadFontStyle
#[excel(content_font_style(...))]      // @ContentFontStyle
#[excel(once_absolute_merge(...))]     // @OnceAbsoluteMerge
struct Demo {
    #[excel(name = "Name", index = 0)] // @ExcelProperty
    name: String,

    #[excel(ignore)]                    // @ExcelIgnore
    internal: String,

    #[excel(format = "yyyy-MM-dd")]     // @DateTimeFormat
    date: chrono::NaiveDate,

    #[excel(column_width = 30)]         // @ColumnWidth (field-level)
    #[excel(content_loop_merge(each_row = 2, column_extend = 1))]
    data: String,

    #[excel(converter = MyConverter)]   // @ExcelProperty.converter
    custom: String,
}
```

## Handler Lifecycle

Write handlers follow Java's event order:

```
before_workbook → after_workbook
    ├── before_sheet → after_sheet
    │       ├── before_row → after_row
    │       │       ├── before_cell → after_cell
    │       │       │       └── (style_cell_style / style_column_width / ...)
    │       │       └── ...
    │       └── ...
    └── finish / finish_on_exception
```

## Test Statistics

| Category | Count | Status |
|----------|-------|--------|
| Total tests | 1315+ | All pass |
| Golden tests (Java output comparison) | 88 | All pass |
| Parity tests (behavioral equivalence) | 152 | All pass |
| 1:1 method tests | 78 | All pass |
| `#[ignore]` annotations | 0 | Eliminated |
