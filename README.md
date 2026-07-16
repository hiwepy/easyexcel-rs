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
