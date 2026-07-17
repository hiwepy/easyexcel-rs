# easyexcel-rs

`easyexcel-rs` aims to provide the Java Alibaba EasyExcel programming model in
idiomatic Rust: builders, typed row mapping, event listeners, converters,
streaming reads, constant-memory writes, templates, and write handlers.

The project is under active compatibility development. The authoritative
feature inventory is [docs/compatibility.md](docs/compatibility.md).

```rust,no_run
use easyexcel::{EasyExcel, ExcelRow, PageReadListener};

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
For Java `OutputStream`-style integration, use the re-exported
`write_csv_to_writer` function with any owned Rust `Write` implementation and a
logical path for handler context.

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

Legacy `.xls` files use the same typed read builders and listener lifecycle.
The binary worksheet is materialized by Calamine before dispatch; writing
`.xls` returns an explicit unsupported-operation error instead of emitting an
XLSX package with the wrong extension.

XLSX worksheet cells are read incrementally through Calamine's `quick-xml`
cell stream rather than materializing the worksheet. Listener callbacks follow
Java EasyExcel's ordering, including workbook-wide `has_next` termination and
exception routing. With `ignore_empty_row(false)`, a row-metadata-only OOXML
scan also preserves leading, intermediate, and trailing empty-row callbacks;
the default path skips that extra scan. The remaining SAX compatibility work is
tracked in [the compatibility contract](docs/compatibility.md#xlsx-streaming-boundary).
Shared and inline rich strings, booleans, numbers, cached formula results,
error text, and 1900/1904 dates follow Java's typed read behavior. String cells
and headers are trimmed by default; call `.auto_trim(false)` to preserve their
outer whitespace. The same option controls whitespace-tolerant sheet-name
matching, using Java `String.trim()` semantics.

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
