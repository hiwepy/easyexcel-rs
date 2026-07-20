# Java golden expectations

JSON snapshots produced by **true Java EasyExcel 4.0.3** for Rust对照.

## Generate / refresh

```bash
# Requires JDK 8+ and Maven 3.6+
./scripts/export-java-golden.sh

# Optional overrides
EASYEXCEL_JAVA_HOME=/path/to/jdk ./scripts/export-java-golden.sh
FIXTURES_DIR=/path/to/fixtures OUT_DIR=/path/to/golden ./scripts/export-java-golden.sh
```

Implementation: `scripts/java-golden-export` (Maven `exec:java` →
`com.alibaba.easyexcel.golden.JavaGoldenExporter`).

Committed `*.expected.json` (and `artifacts/`) must be enough for
`cargo test -p easyexcel --test java_golden_tests` **without** a local JDK.
Missing goldens **fail** the test (no soft-skip); re-run the script above.

Current count: **≥100** (currently **103**, ofNoRows=0; see `docs/test-parity-status.md`).

## JSON schema

```json
{
  "source": "com.alibaba...#method",
  "fixture": "relative/path.or/artifacts/file",
  "sheet_index": 0,
  "sheet_name": "optional",
  "head_row_number": 0,
  "password": "optional",
  "row_count": 10,
  "cells": { "0.0": "姓名0" },
  "rows": [[ "col0", "col1" ]]
}
```

Cell values are Java `ReadDefaultReturnEnum.STRING` display text.
All checked-in goldens currently include full `rows` (ofNoRows cleared).
If `rows` were empty, only `row_count` + `cells` would be asserted (reserved for known format gaps).

## Coverage (summary)

| Area | Examples |
|------|----------|
| Compatibility | t01.xls … t07, t09 |
| BOM / demo | office_bom, demo.xlsx/csv, extra, cellData |
| Converter | converter07/03/csv + write artifact |
| Dataformat | xlsx/xls/v2/date1/date2 |
| Multi-sheet | xlsx + xls sheet0/1 |
| Write artifacts | simple, fill(+horizontal/byName), style, exclude/include, no-head, sort, encrypt |
| Core classes | cache / celldata / charset / exception / handler / large-sample / nomodel / list-head |
| Template | template07 / template03.xls read |
