# CodeGraph Gap Audit: Java vs Rust

> Generated 2026-07-21 by direct codegraph_node queries comparing
> `com.alibaba:easyexcel` (Java 4.0.3) with `easyexcel-rs`.
>
> **Goal**: identify every Java method that does NOT have a
> corresponding Rust method with the same name (modulo Rust
> snake_case migration) and the same semantics.

## A. Critical gaps (large method-count deficit)

| Class | Java methods | Rust methods | Deficit | Notes |
|-------|--------------|--------------|---------|-------|
| `ExcelReaderSheetBuilder` | 10 | 3 | -7 | Missing: `build()`, `doRead()`, `doReadSync()`, `parameter()`, plus 3 from AbstractExcelReaderParameterBuilder (`headRowNumber`, `useScientificFormat`, `registerReadListener`) |
| `AbstractExcelReaderParameterBuilder` | 3 | 1 (via builder) | -2 | Missing: `headRowNumber`, `useScientificFormat`, `registerReadListener` on SheetBuilder |
| `ReadSheet` | 11 | 4 | -7 | Missing: `setSheetNo`, `setSheetName`, `copyBasicParameter`, `toString`, plus 3 constructors |
| `ExcelWriterSheetBuilder` | 0 (found) | n/a | OK | Probably exists in Java AbstractExcelWriterParameterBuilder |

## B. Medium gaps (3-5 methods deficit)

| Class | Java | Rust | Notes |
|-------|------|------|-------|
| `ExcelWriterBuilder` | 28 | 32 | +4 (Rust has more variants: `head()`, `to_writer()`) |
| `ExcelReaderBuilder` | 25 | 30 | +5 (Rust has locale, custom_object, etc.) |
| `ExcelReaderTableBuilder` | ~5 | ? | folded into sheet builder in Rust |
| `WriteSheet` | ~5 | ? | TBD |
| `WriteTable` | ~5 | ? | TBD |
| `WriteWorkbook` | ~3 | ? | TBD |
| `ReadWorkbook` | ~3 | ? | TBD |

## C. Small / acceptable gaps

| Item | Notes |
|------|-------|
| POI `Workbook`/`Sheet`/`Cell` POJO types | Folded into context flags (Java CellValue handles display); Rust does not expose raw POI handles |
| `getInputStream()`/`getOriginalInputStream()` | Folded into ReadOptions |
| `removeThreadLocalCache()` / `clearEncrypt03()` | Folded into per-read context lifecycle |
| `WorkbookWriteContext.getWorkbookHolder()` etc. | Folded into WriteContext |
| `finish(boolean onException)` split into `finish()` / `finish_on_exception()` / `finish_with_exception(bool)` | Rust idiom: explicit methods > flag arg |

## D. Identified gaps to fill in Phase E (priority order)

### D1. **ExcelReaderSheetBuilder** — needs full method surface
Add: `parameter()` returning `&ReadSheet`, `build()` returning `ReadSheet`,
`do_read()` and `do_read_sync()` calling back to `ExcelReader`.

### D2. **ReadSheet** — needs setters + copy + toString
Add: `set_sheet_no`, `set_sheet_name`, `copy_basic_parameter`, `to_string`.

### D3. **WriteSheet** — Java has ~5 methods (sheet_no/sheet_name ctors etc.)

### D4. **WriteTable** — Java has 5+ methods (ctors + setters)

### D5. **WriteWorkbook / ReadWorkbook** — Java has ~5 methods each

### D6. **ExcelReaderTableBuilder** — separate from sheet builder

### D7. **AbstractExcelReaderParameterBuilder** — needs to be implemented
on ExcelReaderSheetBuilder so `headRowNumber` etc. work as in Java.

## E. Handler strategy gaps (acceptable)

These Java strategies already exist in Rust via 1:1 mirroring:

| Java class | Rust type | Status |
|---|---|---|
| `AbstractCellStyleStrategy` | `AbstractCellStyleStrategy` | OK |
| `HorizontalCellStyleStrategy` | `HorizontalCellStyleStrategy` | OK |
| `AbstractVerticalCellStyleStrategy` | `AbstractVerticalCellStyleStrategy` | OK |
| `VerticalCellStyleStrategy` | `VerticalCellStyleStrategy` | OK |
| `AbstractColumnWidthStyleStrategy` | `AbstractColumnWidthStyleStrategy` | OK |
| `SimpleColumnWidthStyleStrategy` | `SimpleColumnWidthStyleStrategy` | OK |
| `LongestMatchColumnWidthStyleStrategy` | `LongestMatchColumnWidthStyleStrategy` | OK |
| `AbstractRowHeightStyleStrategy` | `AbstractRowHeightStyleStrategy` | OK |
| `SimpleRowHeightStyleStrategy` | `SimpleRowHeightStyleStrategy` | OK |
| `LoopMergeStrategy` | `LoopMergeStrategy` | OK |
| `OnceAbsoluteMergeStrategy` | `OnceAbsoluteMergeStrategy` | OK |

## F. Converter completeness

60+ Java converters all have Rust counterparts with matching
file/module names. The Rust versions collapse Java's 4-method
interface (supportJavaTypeKey, supportExcelTypeKey, 2× convertToJavaData,
2× convertToExcelData) into the single `Converter<T>` trait with
generic type parameter + 3 essential methods.

## G. Action plan for Phase E

1. Expand `ExcelReaderSheetBuilder` to full Java surface
2. Expand `ReadSheet` with setters/copy/toString
3. Audit and expand `WriteSheet` / `WriteTable` / `WriteWorkbook`
4. Audit and expand `ReadWorkbook`
5. Implement `ExcelReaderTableBuilder`
6. Move `AbstractExcelReaderParameterBuilder` methods into
   `ExcelReaderSheetBuilder` as direct impl methods
7. Add comprehensive 1:1 tests for each filled-in class
8. Verify all existing 1271 tests still pass

## H. Naming convention audit

| Java | Rust | Form |
|------|------|------|
| camelCase | snake_case | rust idiomatic |
| Abstract* prefix | Abstract* prefix | preserved |
| get*/set* prefix | method (no prefix) | Rust idiomatic |
| *Impl suffix | *Impl suffix | preserved |
| static factory | module fn | rust idiomatic |
| interface | trait | rust idiomatic |
| class with state | struct | rust idiomatic |
| default method | default method | preserved |
| overload (same name, different sigs) | distinct names | Rust idiom: separate methods |

## I. Verification plan

After Phase E, run:
```bash
cargo test --workspace --all-features --no-fail-fast
```
Expected: 1271 + new_tests all pass, 0 failures.

Then run:
```bash
cargo test -p easyexcel --test java_golden_tests
```
Expected: 88/88 golden tests still pass.