# Hutool POI Excel adoption plan

`easyexcel-rs` follows Java EasyExcel as its compatibility contract. Hutool POI
is a secondary source of practical Excel ergonomics that may be added only when
they preserve EasyExcel's event-driven, low-memory model. This document records
the boundary so the future Rust Hutool POI facade can depend on `easyexcel-rs`
without forcing random-access workbook behavior into the core.

## Adoption rules

1. Java EasyExcel naming, defaults, callback order, and file semantics win when
   the two Java projects differ.
2. Hutool-inspired APIs are additive and must not change an EasyExcel default.
3. A feature belongs in `easyexcel-rs` only when it works with streaming rows or
   write-time worksheet metadata.
4. Backend limitations return typed errors; they never trigger silent workbook
   materialization or format substitution.
5. Every adopted behavior requires API tests, backend fixtures, and workspace
   line, function, and region coverage of 100%.

## Capabilities worth adopting

| Hutool POI behavior | Rust direction | Priority | Reason |
|---|---|---:|---|
| `setHeaderAlias` / `addHeaderAlias` | read-builder `header_alias` applied before typed mapping | P0 | Handles external workbooks whose headings cannot be changed |
| inclusive `startRowIndex` / `endRowIndex` | `start_row`, `end_row`, and `read_rows` while retaining header analysis | P0 | Efficient import slicing without collecting unwanted rows |
| `CellEditor` | a typed pre-conversion cell transform with coordinates and source metadata | P0 | Centralizes normalization that otherwise needs many field converters |
| `addSelect` | worksheet data-validation metadata for explicit lists and ranges | P0 | Common business-export requirement and naturally write-time metadata |
| `GlobalPoiConfig` ZIP limits | per-reader file, entry, expanded-byte, row, column, and text limits | P0 | Production guardrail against resource exhaustion and malformed packages |
| `BigExcelWriter` lifecycle | explicit constant-memory policy and measurable memory/disk options | P0 | Reinforces EasyExcel's main large-file promise |
| `writeRow(Map)` and alias ordering | ordered `DynamicRow` writing plus runtime heads and aliases | P1 | Useful for schema-less exports; must retain deterministic column order |
| `FormulaCellValue` | rich formula value with cached value and formula metadata | P1 | Improves round-trip and handler ergonomics |
| merged-cell top-left lookup | optional merged-value propagation during reads | P1 | Useful for report imports; disabled by default to preserve physical cells |
| picture extraction by anchor | streamed XLSX drawing relationship extraction | P1 | Completes a known EasyExcel image-read gap |
| width ratio and tracked auto-size | width scaling with an explicit memory/cost contract | P1 | Better CJK output while keeping large-write behavior predictable |
| row/cell callbacks | optional cell event hook in addition to row listeners | P1 | Enables auditing and fine-grained validation without a workbook model |

The first P0 slice is implemented by `header_alias`, `start_row`, `end_row`,
and `read_rows`. Physical row bounds are inclusive. Header callbacks and header
mapping still run even when the configured data range starts below the header.

## Capabilities reserved for a Hutool-style facade

The following operations require a mutable random-access workbook and therefore
do not belong in the EasyExcel core:

- arbitrary `readCellValue`, `readRow`, and `readColumn` workbook queries;
- inserting or deleting existing rows while repairing merged regions;
- editing an existing workbook in place with full style retention;
- cloning, renaming, reordering, hiding, or activating sheets;
- direct cell/style/font creation utilities and POI-like object access;
- formula recalculation engines and workbook-wide evaluation;
- unrestricted picture, drawing, chart, and shape mutation.

A future Rust Hutool POI implementation can expose these conveniences through a
separate random-access engine and delegate streaming imports/exports to
`easyexcel-rs`. This keeps the dependency direction clear: the Hutool facade may
depend on EasyExcel, while EasyExcel does not depend on the Hutool facade.

## Delivery order

1. Finish Java EasyExcel public API and fixture parity.
2. Add P0 Hutool-inspired features only where they close a production gap and
   do not delay an equivalent EasyExcel feature.
3. Complete image reads, rich formulas, and optional merged-value propagation.
4. Build the separate Hutool-style Rust facade on top of the stable public
   `easyexcel-rs` traits and metadata types.

