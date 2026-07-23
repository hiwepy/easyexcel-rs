# easyexcel-rs Migration Status Tracker

> 迁移进度追踪。以 `docs/migration/file-map.csv` + `xtask migration-audit` 为准。
> **不以「测试绿」替代账本 `complete`。**

## 0. Baseline（Phase 0，2026-07-23）

| Metric | Value |
|--------|-------|
| Java 基线 | EasyExcel **4.0.3** @ `3afdea9d` |
| Rust HEAD | 工作树（本轮结构对齐） |
| Java main 文件（excl package-info） | **325** |
| file-map 行数 | **325** |
| rust 目标文件缺失 | **0**（含 1:1 路径 shim / re-export） |
| Workspace | `crates/*` + `easyexcel-web/{axum,actix}` + `easyexcel-demo/*` + `xtask` |
| JSON | Jackson/Fastjson → `serde` / `serde_json`（`ExcelDownloadErrorBody`） |
| Web | Spring Boot → `easyexcel-web-axum`；Quarkus → `easyexcel-web-actix` |

## 1. 本轮已完成

- [x] 生成 `docs/migration/file-map.csv`
- [x] 落地 `xtask`（`migration-audit` / `migration-audit-strict`）
- [x] 补齐 37 个「路径级」缺失文件（不删既有 `enum_*.rs` / Builder 实现，仅 shim + 新类型）
- [x] 新增 `Font` / `CellData` / `DataFormatData` / `ReadBasicParameter` / `Empty`
- [x] `easyexcel-web-axum` + `easyexcel-web-actix` + 六个 demo
- [x] 修复 `enums.rs`/`enums/mod.rs`、`support.rs`/`support/mod.rs` 双模块冲突（现代 `foo.rs + foo/`）
- [x] `cargo check --workspace` 通过

## 2. 待继续（不得删减既有实现）

- [ ] 将 `in_progress` 账本项补齐 `test_evidence` 后升为 `complete`
- [ ] 逐步把残留 `mod.rs` 迁移为 `foo.rs + foo/`（66 处，渐进）
- [ ] Java `easyexcel-test` 全量用例 → Rust parity/golden（Phase E）
- [ ] `migration-audit-strict` 全绿（Phase G / v1.0）

## 3. 验收命令

```bash
cargo run -p xtask -- migration-audit
cargo check --workspace --all-targets
cargo test -p easyexcel-core excel_download_error_body -- --nocapture
cargo test -p easyexcel-web-axum -- --nocapture
cargo test -p easyexcel-web-actix -- --nocapture
```
