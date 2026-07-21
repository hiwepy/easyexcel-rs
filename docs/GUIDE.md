# easyexcel-rs Usage Guide

This guide covers the most common patterns for reading, writing, and
filling Excel files with easyexcel-rs.

## Reading

### Basic Typed Read

```rust
use easyexcel::{EasyExcel, ExcelRow};

#[derive(Debug, ExcelRow)]
struct User {
    #[excel(name = "Name", index = 0)]
    name: String,
    #[excel(name = "Age", index = 1)]
    age: u32,
}

let users: Vec<User> = EasyExcel::read_sync::<User>("users.xlsx")
    .sheet("Users")
    .do_read_sync()?;
```

### Event-driven Read (Large Files)

```rust
use easyexcel::{EasyExcel, ExcelRow, ReadListener, AnalysisContext};

struct PrintListener;

impl ReadListener<User> for PrintListener {
    fn invoke(&mut self, data: User, _ctx: &AnalysisContext) -> easyexcel::Result<()> {
        println!("{:?}", data);
        Ok(())
    }
}

EasyExcel::read::<User, _>("large.xlsx", PrintListener)
    .ignore_empty_row(true)
    .auto_trim(true)
    .do_read()?;
```

### Paged Read

```rust
use easyexcel::PageReadListener;

let listener = PageReadListener::new(500, |rows: Vec<User>, _ctx| {
    db::batch_insert(rows);
});
EasyExcel::read::<User, _>("data.xlsx", listener).do_read()?;
```

### Dynamic (No-Model) Read

```rust
let rows = EasyExcel::read_dynamic_sync("data.xlsx")
    .head_row_number(1)
    .do_read_sync()?;
// rows: Vec<DynamicRow>
// row.values() -> (col_index, CellValue)
```

### Reading Extras (Comments / Hyperlinks)

```rust
EasyExcel::read::<User, _>("data.xlsx", listener)
    .extra_read(CellExtraType::Comment)
    .extra_read(CellExtraType::Hyperlink)
    .do_read()?;

// listener.extra() will be called for each comment/hyperlink
```

### Password-Protected Read

```rust
EasyExcel::read::<User, _>("encrypted.xlsx", listener)
    .password("123456")
    .do_read()?;
```

## Writing

### Basic Write

```rust
EasyExcel::write::<User>("output.xlsx")
    .sheet("Users")
    .do_write(users)?;
```

### Streaming / Constant Memory Write

```rust
EasyExcel::write::<User>("large_output.xlsx")
    .constant_memory(true)
    .compress_temp_files(true) // SXSSF gzip spill
    .sheet("Data")
    .do_write_iter(users.into_iter())?;
```

### Password-Protected Write

```rust
EasyExcel::write::<User>("encrypted.xlsx")
    .password("123456")
    .sheet("Data")
    .do_write(users)?;
```

### With Styles and Formatting

```rust
use easyexcel::ExcelRow;

#[derive(Debug, ExcelRow)]
#[excel(column_width = 20)]
struct StyledUser {
    #[excel(name = "Name", column_width = 30)]
    #[excel(head_font_style(bold = true))]
    name: String,
    #[excel(name = "Birthday", format = "yyyy-MM-dd")]
    birthday: chrono::NaiveDate,
    #[excel(name = "Score")]
    #[excel(content_style(fill_foreground_color = "green"))]
    score: u32,
}
```

### With Merge Cells

```rust
#[derive(Debug, ExcelRow)]
#[excel(once_absolute_merge(first_row_index = 0, last_row_index = 0, first_column_index = 0, last_column_index = 2))]
struct MergedTitle {
    #[excel(name = "Title")]
    title: String,
    #[excel(content_loop_merge(each_row = 3, column_extend = 1))]
    repeated: String,
}
```

### Exclude / Include Columns

```rust
EasyExcel::write::<User>("output.xlsx")
    .exclude_column_field_names(["internal_id"]) // skip by field name
    .include_column_indexes([0, 2])              // only col 0 and 2
    .sheet("Users")
    .do_write(users)?;
```

### Multi-Sheet Write

```rust
let mut writer = EasyExcel::write::<User>("output.xlsx").build();
writer.write(users, &EasyExcel::writer_sheet::<User>("Sheet1"))?;
writer.write(admins, &EasyExcel::writer_sheet::<User>("Admins"))?;
writer.finish()?;
```

### Write with Image

```rust
let writer = EasyExcel::write::<User>("output.xls");
// For XLS: embed raw image bytes
writer.write_image(&std::fs::read("logo.jpg")?, 0, 5);
writer.sheet("Data").do_write(users)?;
```

### CSV with BOM

```rust
EasyExcel::write::<User>("output.csv")
    .with_bom(true)
    .charset("GBK")
    .sheet("Data")
    .do_write(users)?;
```

## Template Fill

### Scalar Fill ({key})

```rust
let data = TemplateData::new()
    .with("company", "Acme Corp")
    .with("date", "2024-01-15")
    .with("total", 1500);
EasyExcel::fill_template("invoice_template.xlsx", "invoice.xlsx", &data)?;
```

### List Fill ({.field})

```rust
let items = FillWrapper::new([
    TemplateData::new().with("name", "Widget A").with("qty", 10).with("price", 15.0),
    TemplateData::new().with("name", "Widget B").with("qty", 5).with("price", 42.0),
]);
EasyExcel::fill_template_list("template.xlsx", "output.xlsx", &items, FillConfig::default())?;
```

### Named List Fill ({prefix.field})

```rust
let items = FillWrapper::named("items", [
    TemplateData::new().with("name", "A").with("price", 10.0),
]);
// Template contains {items.name}, {items.price}
EasyExcel::fill_template_list("template.xlsx", "output.xlsx", &items, FillConfig::default())?;
```

### Horizontal Fill

```rust
use easyexcel::FillDirection;

let items = FillWrapper::new([...]);
let config = FillConfig::new().direction(FillDirection::Horizontal);
EasyExcel::fill_template_list("template.xlsx", "output.xlsx", &items, config)?;
```

### Force New Row

```rust
let config = FillConfig::new().force_new_row(true);
// Each data row creates a new physical row (default reuses template row)
```

## Converters

### Use Built-in Converter

```rust
#[derive(Debug, ExcelRow)]
struct Data {
    #[excel(converter = MyDateConverter)]
    date: String, // custom format handled by converter
}
```

### Register Global Converter

```rust
EasyExcel::read::<User, _>("data.xlsx", listener)
    .register_converter::<String, YesNoConverter>(YesNoConverter)
    .do_read()?;
```

## Error Handling

```rust
use easyexcel::ExcelError;

match EasyExcel::read_sync::<User>("data.xlsx").do_read_sync() {
    Ok(users) => { /* success */ }
    Err(ExcelError::SheetNotFound(name)) => eprintln!("Sheet '{}' not found", name),
    Err(ExcelError::Format(msg)) => eprintln!("Format error: {}", msg),
    Err(ExcelError::Io(e)) => eprintln!("I/O error: {}", e),
    Err(e) => eprintln!("Other: {}", e),
}
```

## More Examples

See `crates/easyexcel/tests/` for comprehensive 1:1 Java test parity examples covering
every API path: annotation combinations, converters, encrypt, fill, handlers, large data,
multi-sheet, styles, and more.