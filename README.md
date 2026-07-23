# easyexcel-rust

[![Rust](https://img.shields.io/badge/rust-1.88+-blue.svg)](https://www.rust-lang.org)
[![License](https://img.shields.io/badge/license-Apache--2.0-green.svg)](LICENSE)
[![tests](https://img.shields.io/badge/tests-1315+-green.svg)](https://github.com/easy-4-rust/easyexcel-rust)

**easyexcel-rust** is a native Rust port of Alibaba [EasyExcel](https://github.com/alibaba/easyexcel) 4.0.3.
It delivers the Java EasyExcel programming model in idiomatic Rust: builders,
typed row mapping, event listeners, converters, streaming reads,
constant-memory writes, template filling, and write handlers.

> 📖 [中文 README](README_CN.md) | [Architecture](docs/ARCHITECTURE.md) | [Usage Guide](docs/GUIDE.md)

---

## Quick Start

```toml
[dependencies]
easyexcel = "0.1"
```

### Read Excel

```rust
use easyexcel::{EasyExcel, ExcelRow, PageReadListener};

#[derive(Debug, ExcelRow)]
struct User {
    #[excel(name = "Name", index = 0)]
    name: String,
    #[excel(name = "Age", index = 1)]
    age: Option<u32>,
}

fn main() -> easyexcel::Result<()> {
    // Event-driven for large files
    let listener = PageReadListener::new(1000, |rows, _ctx| {
        println!("received {} rows", rows.len());
    });
    EasyExcel::read::<User, _>("users.xlsx", listener)
        .sheet("Users")
        .do_read()?;

    // Synchronous for small datasets
    let users: Vec<User> = EasyExcel::read_sync::<User>("users.xlsx")
        .sheet("Users")
        .do_read_sync()?;

    Ok(())
}
```

### Write Excel

```rust
use easyexcel::{EasyExcel, ExcelRow};

#[derive(Debug, ExcelRow)]
#[excel(column_width = 18)]
struct User {
    #[excel(name = "Name", column_width = 30)]
    name: String,
    #[excel(name = "Age")]
    age: u32,
    #[excel(name = "Birthday", format = "yyyy-MM-dd")]
    birthday: chrono::NaiveDate,
}

fn main() -> easyexcel::Result<()> {
    let users = vec![
        User { name: "Alice".into(), age: 28, birthday: chrono::NaiveDate::from_ymd_opt(1996, 5, 20).unwrap() },
        User { name: "Bob".into(), age: 32, birthday: chrono::NaiveDate::from_ymd_opt(1992, 3, 15).unwrap() },
    ];

    EasyExcel::write::<User>("users.xlsx")
        .sheet("Users")
        .do_write(users)?;

    Ok(())
}
```

### Template Fill

```rust
use easyexcel::{EasyExcel, TemplateData, FillWrapper, FillConfig};

// Scalar fill {key}
let data = TemplateData::new()
    .with("name", "Alice")
    .with("date", "2024-01-15");
EasyExcel::fill_template("template.xlsx", "output.xlsx", &data)?;

// List fill {.field}
let list = FillWrapper::new([
    TemplateData::new().with("name", "Alice").with("score", 95),
    TemplateData::new().with("name", "Bob").with("score", 88),
]);
EasyExcel::fill_template_list("template.xlsx", "output.xlsx", &list, FillConfig::default())?;
```

## Annotation Mapping (Java → Rust)

| Java Annotation | Rust Attribute | Purpose |
|-----------------|---------------|---------|
| `@ExcelProperty` | `#[excel(name, index, order, converter)]` | Column mapping |
| `@ExcelIgnore` | `#[excel(ignore)]` | Skip field |
| `@ExcelIgnoreUnannotated` | `#[excel(ignore_unannotated)]` | Skip unannotated |
| `@DateTimeFormat` | `#[excel(format = "...")]` | Date format |
| `@NumberFormat` | `#[excel(format = "...")]` | Numeric format |
| `@ColumnWidth` | `#[excel(column_width = N)]` | Column width |
| `@HeadRowHeight` | `#[excel(head_row_height = N)]` | Header row height |
| `@ContentRowHeight` | `#[excel(content_row_height = N)]` | Content row height |
| `@HeadStyle` | `#[excel(head_style(...))]` | Header style |
| `@ContentStyle` | `#[excel(content_style(...))]` | Content style |
| `@HeadFontStyle` | `#[excel(head_font_style(...))]` | Header font |
| `@ContentFontStyle` | `#[excel(content_font_style(...))]` | Content font |
| `@ContentLoopMerge` | `#[excel(content_loop_merge(...))]` | Loop merge |
| `@OnceAbsoluteMerge` | `#[excel(once_absolute_merge(...))]` | Absolute merge |

## Write Handlers

```rust
use easyexcel::WriteHandler;
use easyexcel_core::{WriteSheetContext, Result};

struct LoggingHandler;

impl WriteHandler for LoggingHandler {
    fn order(&self) -> i32 { 100 }
    fn after_sheet(&mut self, ctx: &WriteSheetContext) -> Result<()> {
        println!("Sheet '{}' written", ctx.sheet_name());
        Ok(())
    }
}

EasyExcel::write::<User>("output.xlsx")
    .register_write_handler(LoggingHandler)
    .sheet("Sheet1")
    .do_write(data)?;
```

## Crate Map

| Crate | Purpose | Java Mirror |
|-------|---------|-------------|
| `easyexcel` | User-facing facade | `EasyExcel` / `EasyExcelFactory` |
| `easyexcel-core` | Traits, data models, errors | `com.alibaba.excel.*` |
| `easyexcel-derive` | `#[derive(ExcelRow)]` proc-macro | Annotation processing |
| `easyexcel-reader` | XLSX/XLS/CSV read engine | `analysis/` + `read/` |
| `easyexcel-writer` | XLSX/XLS/CSV write engine | `write/` |
| `easyexcel-template` | Template fill engine | `write/metadata/fill/` |

## Java Compatibility

easyexcel-rust is a 1:1 mirror of Java EasyExcel 4.0.3:

- **335 Java @Test methods** — all have Rust `#[test]` counterparts
- **88 Golden tests** — byte-level output matches Java
- **152 Parity tests** — end-to-end behavioral equivalence
- **0 FAILEDs** across entire workspace

See [Migration Audit](docs/migration/TEST_AUDIT_REPORT.md).

## License

Apache-2.0