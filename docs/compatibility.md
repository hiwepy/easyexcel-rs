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
| custom `Converter<T>` | `#[excel(converter = Type)]` + converter contexts | partial: field converter implemented |
| `EasyExcel.write(file, head)` | `EasyExcel::write::<T>(file)` | implemented |
| `sheet(Integer/String)` | `sheet_index(index)` / `sheet(name)` | implemented |
| `needHead` | `need_head` | implemented |
| freeze panes | `freeze_head` / `freeze_panes` | implemented |
| `doWrite(Collection)` | `do_write(IntoIterator)` | implemented |
| streaming write | `do_write_iter(IntoIterator)` | implemented |
| `ExcelWriter` multi-write lifecycle | `ExcelWriter::write` / `finish` | implemented |
| `WriteHandler` lifecycle | ordered `WriteHandler` callbacks | implemented |
| include/exclude columns | builder filters | implemented |
| column width / auto width | `column_width` / `auto_width` | implemented |
| `@ColumnWidth` / `@HeadRowHeight` / `@ContentRowHeight` | `#[excel(column_width, head_row_height, content_row_height)]` | implemented: field width overrides type width; explicit builder width overrides annotations |
| `HorizontalCellStyleStrategy` | header and cycling content `CellStyle` | implemented |
| `@HeadStyle` / `@ContentStyle` / `@HeadFontStyle` / `@ContentFontStyle` | `#[excel(head_style(...), content_style(...), head_font_style(...), content_font_style(...))]` | partial: XLSX cell/font metadata, field-over-type replacement, independent cell/font inheritance, explicit-style precedence, and OOXML verification implemented; Rust attributes use RGB colors while Java indexed-color numeric compatibility and Java fixture parity remain |
| formulas/images/comments/hyperlinks | rich `CellValue` variants | partial: XLSX write implemented |
| `OnceAbsoluteMergeStrategy` | `MergeRange` / `merge_cells` | implemented |
| `LoopMergeStrategy` | repeating data-row merge metadata | implemented |
| dynamic and multi-level heads | `head(Vec<Vec<String>>)` | implemented |
| template `fill` | OOXML-preserving template engine | partial: scalar, named/unnamed vertical and horizontal collections, row reuse, `forceNewRow`, `autoStyle`, formula/range metadata shifting implemented |
| CSV read/write | extension-based CSV engine dispatch | partial: typed read/write, headers, column filters, listeners, write handlers, flexible rows, Java-style `charset`/`withBom`, stateful same-sheet multi-write, UTF-8/UTF-16/GBK streaming transcoding, official Java BOM fixtures, and case-insensitive `.csv` dispatch implemented; JVM-only charset providers remain |
| XLS read | calamine BIFF/XLS engine | implemented: sheet selection, typed mapping, listeners, headers, coordinates, multi-sheet Java fixture; worksheet data is materialized in memory |
| XLS write | backend capability guard | unsupported: returns a typed error instead of silently writing XLSX bytes |
| XLSX password/encryption | `password` on read/write builders | partial: ECMA-376 Agile AES-256/SHA-512 write and Agile/Standard OOXML read implemented; correct, wrong, and missing-password paths tested; encrypted binary XLS is unsupported |
| Axum/Actix adapters | `easyexcel-web` | planned |

## Verification evidence required for 1.0

1. Every Java demo and core test fixture has a Rust counterpart.
2. Read event traces match after deterministic normalization.
3. Written workbook OOXML semantics match after normalization.
4. Excel and LibreOffice open all generated fixtures without repair warnings.
5. Million-row read/write benchmarks record time, peak RSS, and temporary disk.
6. `cargo llvm-cov` reports 100% lines, regions, and functions.
7. Formatting, Clippy, tests, docs, MSRV, and security audit are green in CI.

## Encryption boundary

Password-protected `.xlsx` files are OOXML packages stored inside an OLE/CFB
encryption container. The writer emits ECMA-376 Agile Encryption using AES-256
and SHA-512. The reader recognizes the CFB signature and decrypts Agile or
Standard OOXML before handing the package to the normal typed XLSX engine.
Supplying a password for an unencrypted XLSX is harmless, matching Java
EasyExcel's builder behavior.

Encryption currently buffers the plaintext OOXML package in memory before it
is encrypted or parsed. Legacy encrypted `.xls` uses a different BIFF/RC4
mechanism and returns a typed unsupported-format error; it is not covered by
the OOXML password implementation.

## CSV charset boundary

CSV builders accept Java-style, case-insensitive charset labels through
`charset(...)`, and writers expose `with_bom(...)` with Java EasyExcel's
default of `true`. UTF-8 and UTF-16 byte-order marks are emitted only for
encodings that define one; GBK has no BOM. Reading removes a matching BOM and
transcodes incrementally to UTF-8 before typed row conversion.

The built-in backend covers UTF-8, UTF-16LE/BE, GBK, and WHATWG encoding labels
provided by `encoding_rs`. A JVM installation can expose additional custom
`CharsetProvider` implementations; those provider-specific names currently
return a typed unsupported-operation error.

`write_csv_to_writer` accepts any owned `std::io::Write` sink and preserves the
same handler lifecycle as file output. Its logical path is context metadata,
which mirrors Java EasyExcel's `OutputStream` use without requiring a real file.

Stateful `ExcelWriter::write` caches the first `WriteSheet` configuration,
continues row and content-style indexes across batches, and emits the head only
once. XLSX supports repeated writes to the same sheet and multiple sheets; CSV
supports repeated writes to one logical sheet, matching Java's `CsvWorkbook`
single-sheet constraint.
