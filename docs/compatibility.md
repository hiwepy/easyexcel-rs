# Java EasyExcel compatibility contract

This document is the release gate, not a marketing checklist. A row is marked
`implemented` only when it has Rust API tests and Java-compatible fixture tests.

## Compatibility definition

- Public concepts and lifecycle match Java EasyExcel.
- Rust naming uses `snake_case`; Java `doRead` maps to Rust `do_read`.
- Rust generics replace Java `Class<T>`.
- Generated XLSX bytes need not be identical, but workbook semantics must match.
- Unsupported backend behavior must return a typed error, never silently degrade.

## API inventory

| Java surface | Rust surface | Status |
|---|---|---|
| `EasyExcel.read(file, head, listener)` | `EasyExcel::read::<T, _>(file, listener)` | implemented |
| `headRowNumber` | `head_row_number` | implemented |
| `ignoreEmptyRow` | `ignore_empty_row` | implemented |
| `sheet(Integer/String)` | `sheet(index/name)` | implemented |
| `doRead` | `do_read` | implemented |
| `doReadSync` | `read_sync(...).do_read_sync()` | implemented |
| `ReadListener` | `ReadListener<T>` | implemented |
| `PageReadListener` | `PageReadListener<T>` | implemented |
| `AnalysisContext` | `AnalysisContext` | implemented |
| `@ExcelProperty` | `#[excel(name, index, order, format)]` | implemented |
| `@ExcelIgnore` | `#[excel(ignore)]` | implemented |
| `@ExcelIgnoreUnannotated` | `#[excel(ignore_unannotated)]` | implemented |
| built-in scalar converters | `FromExcelCell` / `IntoExcelCell` | partial |
| custom `Converter<T>` | `Converter<T>` | planned |
| `EasyExcel.write(file, head)` | `EasyExcel::write::<T>(file)` | implemented |
| `sheet(Integer/String)` | `sheet(index/name)` | partial: name implemented |
| `needHead` | `need_head` | implemented |
| freeze panes | `freeze_head` / `freeze_panes` | implemented |
| `doWrite(Collection)` | `do_write(IntoIterator)` | implemented |
| streaming write | `do_write_iter(IntoIterator)` | implemented |
| `ExcelWriter` multi-write lifecycle | `ExcelWriter` | planned |
| `WriteHandler` lifecycle | `WriteHandler` | planned |
| include/exclude columns | builder filters | implemented |
| default styles and annotations | style model | planned |
| formulas/images/comments/hyperlinks | `CellValue` extensions | planned |
| merges and dynamic heads | writer metadata | planned |
| template `fill` | template engine | planned |
| CSV read/write | CSV engine | planned |
| XLS read | calamine XLS engine | planned |
| password/encryption | encryption adapter | planned |
| Axum/Actix adapters | `easyexcel-web` | planned |

## Verification evidence required for 1.0

1. Every Java demo and core test fixture has a Rust counterpart.
2. Read event traces match after deterministic normalization.
3. Written workbook OOXML semantics match after normalization.
4. Excel and LibreOffice open all generated fixtures without repair warnings.
5. Million-row read/write benchmarks record time, peak RSS, and temporary disk.
6. `cargo llvm-cov` reports 100% lines, regions, and functions.
7. Formatting, Clippy, tests, docs, MSRV, and security audit are green in CI.
