# easyexcel-rs

`easyexcel-rs` aims to provide the Java Alibaba EasyExcel programming model in
idiomatic Rust: builders, typed row mapping, event listeners, converters,
streaming reads, constant-memory writes, templates, and write handlers.

The project is under active compatibility development. The authoritative
feature inventory is [docs/compatibility.md](docs/compatibility.md).
Hutool POI-inspired additions follow the explicit boundary in
[docs/hutool-poi-adoption.md](docs/hutool-poi-adoption.md): Java EasyExcel
compatibility remains primary, while streaming-safe production conveniences may
be added without changing EasyExcel defaults.
Future `easydoc-rs`, `easyofd-rs`, and `easypdf-rs` work is recorded only in the
[independent ecosystem roadmap](docs/ecosystem-roadmap.md); those formats are
not implemented in this repository.

```rust,no_run
use easyexcel::{CellExtraType, EasyExcel, ExcelRow, PageReadListener};

#[derive(Debug, ExcelRow)]
struct User {
    #[excel(name = "Name", index = 0)]
    name: String,
    #[excel(name = "Age", index = 1)]
    age: Option<u32>,
}

# fn save(_: Vec<User>) -> easyexcel::Result<()> { Ok(()) }
# fn run() -> easyexcel::Result<()> {
let listener = PageReadListener::new(1_000, |rows, _context| save(rows));

EasyExcel::read::<User, _>("users.xlsx", listener)
    .sheet("Users")
    .auto_trim(true)
    .header_alias("用户姓名", "Name")
    .read_rows(1, 100_000)
    .extra_read(CellExtraType::Comment)
    .do_read()?;
# Ok(())
# }
```

Java's `@ColumnWidth`, `@HeadRowHeight`, and `@ContentRowHeight` map to
Rust derive attributes. A field width overrides the type-level default, while
an explicit builder width overrides both:

```rust,no_run
use easyexcel::{EasyExcel, ExcelRow};

#[derive(ExcelRow)]
#[excel(column_width = 18, head_row_height = 24, content_row_height = 16)]
struct User {
    #[excel(name = "姓名", column_width = 30)]
    name: String,
    #[excel(name = "年龄")]
    age: u32,
}

# fn run(users: Vec<User>) -> easyexcel::Result<()> {
EasyExcel::write::<User>("users.xlsx")
    .column_width(1, 22)
    .do_write(users)?;
# Ok(())
# }
```

Java's `@HeadStyle`, `@ContentStyle`, `@HeadFontStyle`, and
`@ContentFontStyle` map to nested derive attributes. Field annotations replace
the corresponding type-level annotation independently, so a field can override
only its head font while inheriting the type-level head cell style. Explicit
builder styles run last:

```rust,no_run
use easyexcel::{EasyExcel, ExcelRow};

#[derive(ExcelRow)]
#[excel(
    head_style(
        horizontal_alignment = "center",
        fill_pattern = "solid",
        fill_foreground_color = 0x00D9_EAF7,
        border_bottom = "thin"
    ),
    head_font_style(font_name = "Arial", font_height_in_points = 12, bold = true),
    content_style(vertical_alignment = "center", wrapped = true),
    content_font_style(color = 0x0033_3333)
)]
struct StyledUser {
    #[excel(
        name = "姓名",
        head_style(fill_foreground_color = 0x00FF_E699),
        head_font_style(bold = false)
    )]
    name: String,
    #[excel(name = "年龄")]
    age: u32,
}

# fn run(users: Vec<StyledUser>) -> easyexcel::Result<()> {
EasyExcel::write::<StyledUser>("styled-users.xlsx")
    .do_write(users)?;
# Ok(())
# }
```

Style color values `0..=64` use Java EasyExcel's Apache POI indexed palette;
larger integers use `0xRRGGBB` RGB. Likewise, `data_format = 14` selects a Java
built-in format index, while `data_format = "0.00"` uses a custom Excel format.

The same read and write builders automatically select the CSV engine for a
`.csv` path. Typed mapping, listeners, column filters, and write handlers keep
the same lifecycle. Like Java EasyExcel, CSV output includes a BOM by default
when the selected charset defines one. Charset names are case-insensitive:

```rust,no_run
use easyexcel::{EasyExcel, ExcelRow};

#[derive(ExcelRow)]
struct User {
    #[excel(name = "姓名")]
    name: String,
}

# fn run(users: Vec<User>) -> easyexcel::Result<()> {
EasyExcel::write::<User>("users.csv")
    .charset("GBK")
    .with_bom(false)
    .do_write(users)?;

let users = EasyExcel::read_sync::<User>("users.csv")
    .charset("gbk")
    .do_read_sync()?;
# let _ = users;
# Ok(())
# }
```

UTF-8, UTF-16LE/BE, GBK, and the other encodings exposed by `encoding_rs` are
transcoded incrementally rather than buffering the complete CSV file.
Java `OutputStream`-style output is available for XLSX and CSV. A borrowed
writer remains caller-owned, which is the Rust equivalent of
`autoCloseStream(false)`:

```rust,no_run
# use easyexcel::{EasyExcel, ExcelRow};
# #[derive(ExcelRow)] struct User { #[excel(name = "Name")] name: String }
# fn run(users: Vec<User>) -> easyexcel::Result<()> {
let mut response = Vec::new();
EasyExcel::write::<User>("users.xlsx")
    .to_writer(&mut response)
    .do_write(users)?;
assert!(response.starts_with(b"PK"));
# Ok(())
# }
```

For stateful multi-batch output, wrap an owned sink in `ExcelOutputStream` and
use `to_output_stream`. Owned streams close by default; `auto_close_stream(false)`
keeps the sink accessible. `finish_on_exception()` discards accumulated output
by default, while `write_excel_on_exception(true)` emits it like Java
EasyExcel. CSV output is staged until finish so the default exception path
does not leak a partial response. The lower-level `write_xlsx_to_writer`,
`write_csv_to_writer`, and `CsvEncodingWriter` APIs remain available when a
builder is not needed.

Java's no-model `Map<Integer, ...>` reads map to the ordered `DynamicRow` type.
The default mode returns strings, while `ActualData` preserves scalar cell
types and exact numeric cells as the re-exported `BigDecimal` type.
`ReadCellData` exposes the exact raw/converted value, Excel-formatted display
text, physical coordinates, and formula metadata. Missing physical columns are
represented by `DynamicValue::Null`, so sparse rows retain Java-compatible
indexes:

```rust,no_run
use easyexcel::{DynamicRow, DynamicValue, EasyExcel, ReadDefaultReturn};

# fn run() -> easyexcel::Result<()> {
let rows = EasyExcel::read_dynamic_sync("users.xlsx")
    .head_row_number(0)
    .read_default_return(ReadDefaultReturn::ActualData)
    .do_read_sync()?;

if let Some(DynamicValue::ActualData(value)) = rows[0].get(1) {
    println!("second physical column: {}", value.as_text());
}

EasyExcel::write::<DynamicRow>("copy.xlsx").do_write(rows)?;
# Ok(())
# }
```

Use `EasyExcel::read_dynamic(path, listener)` for the event-driven equivalent.
Dynamic rows can be written to XLSX or CSV with no synthetic header, or with a
runtime header supplied through `.head(...)`. Index include/exclude filters and
Java-style include ordering apply to dynamic rows as well.

Stateful builders follow Java `ExcelWriter` semantics: repeated writes to the
same sheet append rows without repeating the head. XLSX may target multiple
sheets, while CSV accepts one logical sheet. `writer_sheet_index(index)` and
`WriteSheet::sheet_index(index)` provide Java-style zero-based logical sheet
numbers; a cached number takes precedence over a newly supplied name:

```rust,no_run
# use easyexcel::{EasyExcel, ExcelRow};
# #[derive(ExcelRow)] struct User { #[excel(name = "Name")] name: String }
# fn run(first_page: Vec<User>, second_page: Vec<User>) -> easyexcel::Result<()> {
let sheet = EasyExcel::writer_sheet::<User>("Users");
let mut writer = EasyExcel::write::<User>("users.csv").build();
writer
    .write(first_page, &sheet)?
    .write(second_page, &sheet)?;
writer.finish()?;
# Ok(())
# }
```

Java-style global converters are registered on read or write builders. Read
selection uses the Rust target type plus [`CellDataType`], while writes select
by Rust type. A field-level `#[excel(converter = Type)]` remains the highest
priority override.

```rust,no_run
use easyexcel::{
    CellDataType, CellValue, Converter, EasyExcel, ReadConverterContext,
    WriteConverterContext,
};

struct PrefixConverter;

impl Converter<String> for PrefixConverter {
    fn support_excel_type(&self) -> CellDataType {
        CellDataType::String
    }

    fn convert_to_rust_data(
        &self,
        context: &ReadConverterContext<'_>,
    ) -> easyexcel::Result<String> {
        Ok(format!("custom:{}", context.cell().map_or_else(String::new, CellValue::as_text)))
    }

    fn convert_to_excel_data(
        &self,
        context: &WriteConverterContext<'_, String>,
    ) -> easyexcel::Result<CellValue> {
        Ok(CellValue::String(format!("custom:{}", context.value())))
    }
}

# fn build<T: easyexcel::ExcelRow>(rows: Vec<T>) -> easyexcel::Result<()> {
EasyExcel::write::<T>("users.xlsx")
    .register_converter::<String, _>(PrefixConverter)
    .do_write(rows)?;
# Ok(())
# }
```

Legacy `.xls` files use the same typed read builders and listener lifecycle.
The binary worksheet is materialized by Calamine before dispatch; writing
`.xls` returns an explicit unsupported-operation error instead of emitting an
XLSX package with the wrong extension.

XLSX worksheet cells and `sharedStrings.xml` are read incrementally by the
`quick-xml` OOXML engine rather than materializing the worksheet. Listener
callbacks follow Java EasyExcel's ordering, including workbook-wide `has_next`
termination and exception routing. With `ignore_empty_row(false)`, a
row-metadata-only OOXML scan also preserves leading, intermediate, and trailing
empty-row callbacks;
the default path skips that extra scan. Shared strings use Java EasyExcel's
5,000,000-byte selection boundary: smaller XML parts stay in memory and larger
parts spill to a temporary disk cache. The strategy can be forced when needed:

```rust
# use easyexcel::{DynamicRow, EasyExcel, ReadCacheMode};
# fn example() -> easyexcel::Result<()> {
let rows = EasyExcel::read_sync::<DynamicRow>("users.xlsx")
    .read_cache(ReadCacheMode::Disk)
    .do_read_sync()?;
# let _ = rows;
# Ok(())
# }
```

The remaining SAX compatibility work is
tracked in [the compatibility contract](docs/compatibility.md#xlsx-streaming-boundary).
Shared and inline rich strings, booleans, numbers, cached formula results,
error text, and 1900/1904 dates follow Java's typed read behavior. String cells
and headers are trimmed by default; call `.auto_trim(false)` to preserve their
outer whitespace. The same option controls whitespace-tolerant sheet-name
matching, using Java `String.trim()` semantics.

Schema-less XLSX reads use the same physical cell event to reproduce Java's
formatted `STRING`, exact `ACTUAL_DATA`, and `READ_CELL_DATA` behavior. The
stream resolves built-in/custom number
formats and the 1900/1904 date system, normalizes Excel's 15-significant-digit
numeric representation, and releases each cell immediately. Malformed input is
reported as a typed format error.

For formula cells, typed fields receive Excel's cached result. The original
expression remains available separately through `RowData::formula()` and
`ReadConverterContext::formula()`, mirroring Java `ReadCellData.formulaData`
without changing ordinary scalar conversion.

Java's default image write converters map naturally to Rust field types:
`Vec<u8>`, `Box<[u8]>`, fixed byte arrays, and `PathBuf` write real XLSX
drawing/media parts. `ImageInputStream<R>` consumes the remaining bytes from a
caller-owned `Read` value without closing it, while the re-exported `Url` type
downloads HTTP/HTTPS image bytes with Java's one-second connect and five-second
read defaults. String paths can opt in explicitly with
`#[excel(converter = easyexcel::StringImageConverter)]`, matching Java's
annotation-selected `StringImageConverter`; unreadable files return an error
instead of writing a path as cell text.

`WriteCellData` preserves an ordinary scalar together with Java-compatible
`imageDataList` entries. Each `ImageData` can use absolute or relative first/last
cell coordinates, pixel margins, and all four POI anchor modes; multiple images
are emitted in list order as independent XLSX drawing anchors:

```rust,no_run
use easyexcel::{
    AnchorType, CellValue, ClientAnchorData, CoordinateData, ImageData, WriteCellData,
};

let anchor = ClientAnchorData::new()
    .coordinates(CoordinateData::new().relative_last_column_index(1))
    .left(4)
    .right(4)
    .anchor_type(AnchorType::MoveAndResize);
let png_bytes = std::fs::read("logo.png")?;
let cell = WriteCellData::new(CellValue::String("product".to_owned()))
    .image(ImageData::new(png_bytes).anchor(anchor));
# Ok::<(), std::io::Error>(())
```

`RichTextStringData` mirrors Java's whole-string `WriteFont` plus ordered
`IntervalFont` overrides. Range indices use Java UTF-16 code units, so Chinese,
emoji, overlapping ranges, superscript/subscript, underline, color, font name,
size, charset, bold, italic, and strike-through retain Java semantics. XLSX
stores real rich-text runs; typed reads return the plain text with no invented
format metadata.

Java `extraRead` maps to `.extra_read(CellExtraType::...)`. Enable `Comment`,
`Hyperlink`, or `Merge` on an XLSX reader and implement `ReadListener::extra`
to receive a `CellExtra` with optional text and zero-based first/last row and
column coordinates. Extra callbacks run after row callbacks and before the
sheet completion callback, with the same `on_exception` and `has_next`
control flow as Java EasyExcel. XLS and CSV return a typed unsupported error
when extra metadata is requested.

Run `./scripts/benchmark-million-rows.sh` for the release-scale streaming
benchmark. It writes and reads one million typed rows and reports elapsed time,
peak RSS, and XLSX size; pass a smaller row count as the first argument for a
quick smoke run.

Password-protected `.xlsx` files use the same Java-style builder call on both
read and write paths:

```rust,no_run
use easyexcel::{EasyExcel, ExcelRow};

#[derive(Debug, ExcelRow)]
struct SecretRow {
    #[excel(name = "Value", index = 0)]
    value: String,
}

# fn run(rows: Vec<SecretRow>) -> easyexcel::Result<()> {
EasyExcel::write::<SecretRow>("protected.xlsx")
    .password("123456")
    .do_write(rows)?;

let values = EasyExcel::read_sync::<SecretRow>("protected.xlsx")
    .password("123456")
    .do_read_sync()?;
# let _ = values;
# Ok(())
# }
```

The writer emits ECMA-376 Agile Encryption (AES-256/SHA-512), while the reader
accepts Agile and Standard encrypted OOXML. Encryption buffers the plaintext
OOXML package in memory. Password-protected legacy `.xls` is a separate BIFF
encryption format and is currently unsupported.

Scalar placeholders in an existing workbook can be filled without rebuilding
its OOXML package:

```rust,no_run
use easyexcel::{EasyExcel, TemplateData};

# fn run() -> easyexcel::Result<()> {
let data = TemplateData::new()
    .with("name", "Alice")
    .with("count", 3);

EasyExcel::fill_template("template.xlsx", "report.xlsx", &data)?;
# Ok(())
# }
```

Java-style collection placeholders are supported as `{.field}` or
`{prefix.field}`:

```rust,no_run
use easyexcel::{EasyExcel, FillConfig, FillWrapper, TemplateData};

# fn run() -> easyexcel::Result<()> {
let users = FillWrapper::named("users", [
    TemplateData::new().with("name", "Alice"),
    TemplateData::new().with("name", "Bob"),
]);

EasyExcel::fill_template_list(
    "template.xlsx",
    "report.xlsx",
    &users,
    FillConfig::new().force_new_row(true),
)?;
# Ok(())
# }
```

Repeated Java-style `fill` calls on the same workbook use a stateful template
writer. Collection rows continue from the previous call for each prefix, and
scalar values can be supplied before or after collection data:

```rust,no_run
use easyexcel::{CellValue, EasyExcel, FillConfig, FillWrapper, TemplateData};

# fn run() -> easyexcel::Result<()> {
let mut writer = EasyExcel::template_writer("template.xlsx", "report.xlsx")?;
writer
    .fill_list(
        &FillWrapper::named("users", [TemplateData::new().with("name", "Alice")]),
        FillConfig::new(),
    )?
    .fill_list(
        &FillWrapper::named("users", [TemplateData::new().with("name", "Bob")]),
        FillConfig::new(),
    )?
    .fill(&TemplateData::new().with("title", "User report"))?
    .write_rows([vec![
        CellValue::Empty,
        CellValue::Empty,
        CellValue::String("Total: 2".to_owned()),
    ]])?
    .finish()?;
# Ok(())
# }
```

Escaped braces (`\{` and `\}`), repeated named/unnamed vertical collections,
and repeated horizontal collections are supported. A placeholder that occupies
the whole cell preserves Boolean, integer, floating-point, decimal, date,
date-time, error, and formula types; placeholders mixed with surrounding text
follow Java and produce a string cell. `write_rows` appends ordinary
`CellValue` rows after the filled template cursor, covering the Java
`complexFillWithTable` pattern.

The implementation is verified against Alibaba EasyExcel's official
`simple.xlsx`, `composite.xlsx`, and `complexFillWithTable.xlsx` fixtures.
Changing `FillConfig` for an already-used prefix, selecting a non-first sheet
for appended rows, typed-model/write-handler composition, and template stream
input/output remain compatibility work in progress.

## Quality gates

```bash
cargo fmt --all -- --check
cargo clippy --workspace --all-targets --all-features -- -D warnings
cargo test --workspace --all-features
./scripts/coverage.sh
```

The coverage script fails unless workspace line, region, and function coverage
are all 100%. It excludes only the stable-Rust `proc_macro::TokenStream` bridge
file; the complete `syn`/`proc_macro2` derive implementation remains inside the
100% gate and macro behavior is also exercised by end-to-end derive tests.
