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
    .do_read()?;
# Ok(())
# }
```

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

Legacy `.xls` files use the same typed read builders and listener lifecycle.
The binary worksheet is materialized by Calamine before dispatch; writing
`.xls` returns an explicit unsupported-operation error instead of emitting an
XLSX package with the wrong extension.

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
