# Easy document ecosystem roadmap

`easyexcel-rust` remains an Excel-only project. The following projects are
planned as independent repositories after the EasyExcel compatibility surface
is stable; none of their implementations belong in this workspace.

| Future project | Scope | Likely Rust foundation | Relationship to `easyexcel-rust` |
|---|---|---|---|
| `easydoc-rs` | DOCX creation, typed document models, templates, tables, and images | `docx-rs`, `quick-xml`, `zip` | Reuse common design conventions only |
| `easyofd-rust` | OFD generation, reading, signing-oriented metadata, and validation | `ofd-rs` plus format-specific XML tooling | Independent format engine and API |
| `easypdf-rs` | PDF generation, merge/split, extraction, forms, and rendering adapters | `lopdf`, `printpdf`, PDFium adapters | Independent format engine and API |

## Sequencing

1. Complete and stabilize Java EasyExcel compatibility in `easyexcel-rust`.
2. Extract only genuinely format-neutral conventions, such as error taxonomy,
   resource-limit configuration, listener control flow, and test policy.
3. Design each new project against its own source-format semantics rather than
   forcing an Excel workbook abstraction onto DOCX, OFD, or PDF.
4. Create, implement, test, version, and publish each project independently.

Shared crates may be considered later only after at least two projects prove
the same abstraction is stable. Until then, no placeholder module or dependency
for these future formats will be added to `easyexcel-rust`.
