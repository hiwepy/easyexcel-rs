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
| `autoTrim` | `auto_trim` | implemented: defaults to `true`; sheet-name matching, content/header strings, and trim-to-empty row handling match Java |
| Hutool POI header aliases and row slices | `header_alias` / `start_row` / `end_row` / `read_rows` | additive enhancement: aliases are applied before typed mapping; physical bounds are inclusive and retain header analysis |
| `sheet(Integer/String)` | `sheet(index/name)` | implemented |
| `doRead` | `do_read` | implemented |
| `doReadSync` | `read_sync(...).do_read_sync()` | implemented |
| no-model `Map<Integer, String/Object/ReadCellData<?>>` | `DynamicRow` + `DynamicValue` | implemented: physical column indexes, sparse gaps, header-tail nulls, XLSX/XLS/CSV event and sync reads, schema-less XLSX/CSV writes, and Java official compatibility fixtures `t01`–`t07`/`t09` are verified |
| `ReadDefaultReturnEnum.STRING/ACTUAL_DATA/READ_CELL_DATA` | `ReadDefaultReturn::{String, ActualData, ReadCellData}` | implemented: Java-compatible default, Excel-formatted strings, exact `BigDecimal` numeric values, formula metadata, display text, coordinates, and raw cell values |
| `ReadListener` | `ReadListener<T>` | implemented |
| `extraRead(CellExtraTypeEnum)` / `ReadListener.extra` | `extra_read(CellExtraType)` / `ReadListener::extra` | implemented for XLSX comments, hyperlinks, and merged ranges; XLS/CSV return a typed unsupported error when requested |
| `PageReadListener` | `PageReadListener<T>` | implemented |
| `AnalysisContext` | `AnalysisContext` | implemented |
| `@ExcelProperty` | `#[excel(name, index, order, format)]` | implemented |
| `@ExcelIgnore` | `#[excel(ignore)]` | implemented |
| `@ExcelIgnoreUnannotated` | `#[excel(ignore_unannotated)]` | implemented |
| built-in scalar converters | `FromExcelCell` / `IntoExcelCell` | partial: strings, booleans, signed/unsigned integers, floats, `BigDecimal`, `Option<T>`, `NaiveDate`, and `NaiveDateTime` are implemented; Java URL/locale/temporal converter inventory remains |
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
| `@HeadStyle` / `@ContentStyle` / `@HeadFontStyle` / `@ContentFontStyle` | `#[excel(head_style(...), content_style(...), head_font_style(...), content_font_style(...))]` | partial: XLSX cell/font metadata, field-over-type replacement, independent cell/font inheritance, explicit-style precedence, POI indexed colors `0..=64`, RGB extensions, built-in/custom data formats, and official Java annotation color expectations are verified; custom HSSF palettes and XLS writing remain |
| formulas/images/comments/hyperlinks | formula metadata, `CellExtra`, and rich write values | partial: XLSX formula and comment/hyperlink/merge reads plus XLSX rich writes implemented; image reads remain |
| `OnceAbsoluteMergeStrategy` | `MergeRange` / `merge_cells` | implemented |
| `LoopMergeStrategy` | repeating data-row merge metadata | implemented |
| dynamic and multi-level heads | `head(Vec<Vec<String>>)` | implemented |
| template `fill` | OOXML-preserving template engine | partial: scalar, named/unnamed vertical and horizontal collections, row reuse, `forceNewRow`, `autoStyle`, formula/range metadata shifting implemented |
| CSV read/write | extension-based CSV engine dispatch | partial: typed read/write, headers, column filters, listeners, write handlers, flexible rows, Java-style `charset`/`withBom`, stateful same-sheet multi-write, UTF-8/UTF-16/GBK streaming transcoding, official Java BOM fixtures, and case-insensitive `.csv` dispatch implemented; JVM-only charset providers remain |
| XLSX SAX read lifecycle | Calamine `worksheet_cells_reader` + typed/dynamic row dispatcher | partial: worksheet cells are streamed through `quick-xml`; every header row, leading/intermediate/trailing empty rows, `autoTrim`, shared/inline rich strings, booleans, 15-significant-digit numbers, built-in/custom display formats, exact decimals, 1900/1904 dates, cached formula results plus `FormulaData`, error text, comment/hyperlink/merged-cell extras, typed and no-model rows, listener exception routing, post-callback `hasNext`, workbook-wide stop, and completion callbacks match Java; disk-backed shared-string caching remains |
| XLS read | calamine BIFF/XLS engine | implemented: sheet selection, typed mapping, listeners, headers, coordinates, multi-sheet Java fixture; worksheet data is materialized in memory |
| XLS write | backend capability guard | unsupported: returns a typed error instead of silently writing XLSX bytes |
| XLSX password/encryption | `password` on read/write builders | partial: ECMA-376 Agile AES-256/SHA-512 write and Agile/Standard OOXML read implemented; correct, wrong, and missing-password paths tested; encrypted binary XLS is unsupported |
| Axum/Actix adapters | `easyexcel-web` | planned |

Hutool POI is used only as a secondary ergonomics and production-hardening
reference. The adoption boundary and dependency direction are documented in
[hutool-poi-adoption.md](hutool-poi-adoption.md).

## Verification evidence required for 1.0

1. Every Java demo and core test fixture has a Rust counterpart.
2. Read event traces match after deterministic normalization.
3. Written workbook OOXML semantics match after normalization.
4. Excel and LibreOffice open all generated fixtures without repair warnings.
5. Million-row read/write benchmarks record time, peak RSS, and temporary disk.
6. `cargo llvm-cov` reports 100% lines, regions, and functions.
7. Formatting, Clippy, tests, docs, MSRV, and security audit are green in CI.

## XLSX streaming boundary

The XLSX reader does not build a worksheet `Range`. It uses Calamine's
`worksheet_cells_reader`, whose implementation incrementally parses worksheet
XML with `quick-xml`, groups only the current row, dispatches it, and releases
that row before reading the next one. The Rust listener sequence follows Java
EasyExcel: each header row invokes `invoke_head`; a successful header or data
callback is followed by `has_next`; `false` stops the complete workbook; and a
stopped sheet does not invoke `do_after_all_analysed`. Conversion and listener
callback failures both pass through `on_exception`.

String cells are trimmed by default, including header strings, matching Java's
`GlobalConfiguration.autoTrim = true`. A string that becomes empty after
trimming participates in Java-compatible empty-row filtering. Name-based sheet
selection trims both the requested and actual name while retaining the actual
workbook name in callbacks. Trimming uses Java `String.trim()` semantics, so
only leading and trailing characters at or below U+0020 are removed. Shared strings
and inline rich-text runs are concatenated, Excel `_xHHHH_` escapes are
decoded, formula cells expose their cached typed value, error cells expose
their literal text, and date values honor the workbook's 1900/1904 windowing.
The original formula expression is retained separately as `FormulaData`, just
like Java `ReadCellData.formulaData`; `RowData::formula` and custom converter
contexts can inspect it without replacing the cached value.

For no-model reads, a second worksheet event stream advances in exact physical
cell lockstep with Calamine. It resolves `styles.xml`, custom and built-in
number formats, and workbook date windowing without retaining prior rows.
`STRING` returns the Java-style display value (including Excel date rounding),
while `ACTUAL_DATA` and `READ_CELL_DATA` expose arbitrary-precision
`BigDecimal` values after Excel's 15-significant-digit normalization. Premature
EOF, extra cells, or coordinate disagreement between the streams fails with a
typed format error instead of silently associating metadata with the wrong
cell.

When `extra_read` enables a category, the companion OOXML scanner also emits
Java-style `CellExtra` values for comments, internal and external hyperlinks,
and merged ranges. Worksheet extras are delivered after row events, comments
follow them, and all extras precede `do_after_all_analysed`. A successful
`extra` callback is followed by `has_next`; failures pass through
`on_exception`. The same parser is used for decrypted in-memory XLSX packages.

Like Java's `RowTagHandler`, the dispatcher synthesizes every missing row before
the first cell-bearing row and between later cell-bearing rows. This preserves
empty header callbacks and `ignore_empty_row(false)` data callbacks without
materializing a sheet. When empty rows are requested, a lightweight companion
scanner resolves package/workbook relationships and streams only `<row r>`
metadata so trailing cell-free rows are preserved too. It supports Transitional
and Strict OOXML relationship namespaces, relative and absolute targets, and
case-insensitive ZIP path lookup. The default `ignore_empty_row(true)` path does
not perform this companion scan.

This is not yet the complete Java `XlsxSaxAnalyser` contract. Calamine loads
the workbook shared-string table in memory, whereas Java can select a
disk-backed cache. Image extraction also remains. Those gaps require focused
extensions around the OOXML stream, not a second implementation of the
already-streaming basic cell path.

## Million-row benchmark

The release benchmark is reproducible and intentionally excluded from normal
test runs:

```shell
./scripts/benchmark-million-rows.sh
```

It writes 1,000,000 typed rows with the constant-memory XLSX writer, reads them
through the event listener without collecting rows, verifies the observed row
count, and reports write/read time plus output size. `/usr/bin/time -l` on
macOS or `/usr/bin/time -v` on Linux records peak RSS for the complete run. A
smaller smoke run can be selected with the first argument, for example
`./scripts/benchmark-million-rows.sh 1000`. The latest measured environment and
results are recorded in [benchmarks.md](benchmarks.md).

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
