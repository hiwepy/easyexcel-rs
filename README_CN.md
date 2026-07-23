# easyexcel-rust

[![Rust](https://img.shields.io/badge/rust-1.88+-blue.svg)](https://www.rust-lang.org)
[![License](https://img.shields.io/badge/license-Apache--2.0-green.svg)](LICENSE)

**easyexcel-rust** 是阿里巴巴 [EasyExcel](https://github.com/alibaba/easyexcel) 的 Rust 原生移植版本。
以惯用 Rust 方式提供 Java EasyExcel 编程模型：Builder 模式、类型化行映射、事件监听器、类型转换器、流式读取、常量内存写入、模板填充和写入处理器。

---

## 快速开始

```toml
[dependencies]
easyexcel = "0.1"
```

### 读取 Excel

```rust
use easyexcel::{EasyExcel, ExcelRow, PageReadListener};

#[derive(Debug, ExcelRow)]
struct User {
    #[excel(name = "姓名", index = 0)]
    name: String,
    #[excel(name = "年龄", index = 1)]
    age: Option<u32>,
}

fn main() -> easyexcel::Result<()> {
    // 事件驱动读取（大数据友好）
    let listener = PageReadListener::new(1000, |rows, _ctx| {
        println!("收到 {} 行", rows.len());
    });
    EasyExcel::read::<User, _>("users.xlsx", listener)
        .sheet("用户表")
        .do_read()?;

    // 同步读取（小数据直接获取）
    let users: Vec<User> = EasyExcel::read_sync::<User>("users.xlsx")
        .sheet("用户表")
        .do_read_sync()?;
    
    Ok(())
}
```

### 写入 Excel

```rust
use easyexcel::{EasyExcel, ExcelRow};

#[derive(Debug, ExcelRow)]
#[excel(column_width = 18)]
struct User {
    #[excel(name = "姓名", column_width = 30)]
    name: String,
    #[excel(name = "年龄")]
    age: u32,
    #[excel(name = "生日", format = "yyyy-MM-dd")]
    birthday: chrono::NaiveDate,
}

fn main() -> easyexcel::Result<()> {
    let users = vec![
        User { name: "张三".into(), age: 28, birthday: chrono::NaiveDate::from_ymd_opt(1996, 5, 20).unwrap() },
        User { name: "李四".into(), age: 32, birthday: chrono::NaiveDate::from_ymd_opt(1992, 3, 15).unwrap() },
    ];

    EasyExcel::write::<User>("users.xlsx")
        .sheet("用户表")
        .do_write(users)?;

    Ok(())
}
```

### 模板填充

```rust
use easyexcel::{EasyExcel, TemplateData};

// 简单填充 {key}
let data = TemplateData::new()
    .with("name", "张三")
    .with("date", "2024-01-15");
EasyExcel::fill_template("template.xlsx", "output.xlsx", &data)?;

// 列表填充 {.field}
let list = FillWrapper::new([
    TemplateData::new().with("name", "张三").with("score", 95),
    TemplateData::new().with("name", "李四").with("score", 88),
]);
EasyExcel::fill_template_list("template.xlsx", "output.xlsx", &list, FillConfig::default())?;
```

---

## 核心特性

| 特性 | 支持格式 | 说明 |
|------|---------|------|
| **类型化读写** | XLSX / XLS / CSV | `#[derive(ExcelRow)]` + 注解属性 |
| **事件监听** | XLSX / XLS / CSV | `PageReadListener` / `ReadListener<T>` |
| **流式读取** | XLSX / XLS | SAX 解析，内存可控 |
| **常量内存写入** | XLSX | `SXSSF` 等价实现 |
| **模板填充** | XLSX / XLS | `{key}` / `{.field}` 占位符 |
| **密码加密** | XLSX / XLS | Agile + RC4 |
| **类型转换器** | 全部 | 60+ 内置转换器 |
| **单元格样式** | XLSX / XLS | 字体/填充/对齐/边框 |
| **合并单元格** | XLSX / XLS | `@OnceAbsoluteMerge` / `@ContentLoopMerge` |
| **批注/超链接** | XLSX | 读+写 |
| **图片** | XLSX | 读+写 |
| **公式** | XLSX | 读+写 |
| **CSV BOM** | CSV | 读写支持 |

---

## 注解映射（Java → Rust）

| Java 注解 | Rust 属性 | 说明 |
|-----------|----------|------|
| `@ExcelProperty` | `#[excel(name, index, order, converter)]` | 列映射 |
| `@ExcelIgnore` | `#[excel(ignore)]` | 忽略字段 |
| `@ExcelIgnoreUnannotated` | `#[excel(ignore_unannotated)]` | 忽略未注解 |
| `@DateTimeFormat` | `#[excel(format = "...")]` | 日期格式 |
| `@NumberFormat` | `#[excel(format = "...")]` | 数字格式 |
| `@ColumnWidth` | `#[excel(column_width = N)]` | 列宽 |
| `@HeadRowHeight` | `#[excel(head_row_height = N)]` | 表头行高 |
| `@ContentRowHeight` | `#[excel(content_row_height = N)]` | 内容行高 |
| `@HeadStyle` | `#[excel(head_style(...))]` | 表头样式 |
| `@ContentStyle` | `#[excel(content_style(...))]` | 内容样式 |
| `@HeadFontStyle` | `#[excel(head_font_style(...))]` | 表头字体 |
| `@ContentFontStyle` | `#[excel(content_font_style(...))]` | 内容字体 |
| `@ContentLoopMerge` | `#[excel(content_loop_merge(...))]` | 循环合并 |
| `@OnceAbsoluteMerge` | `#[excel(once_absolute_merge(...))]` | 绝对合并 |

---

## 写入处理器

```rust
use easyexcel::WriteHandler;
use easyexcel_core::{WriteSheetContext, Result, ExcelCellStyle};

struct MyStyleHandler;

impl WriteHandler for MyStyleHandler {
    fn order(&self) -> i32 { 100 }

    fn after_sheet(&mut self, _ctx: &WriteSheetContext) -> Result<()> {
        println!("Sheet written successfully");
        Ok(())
    }

    fn style_cell_style(&self, _ctx: &easyexcel_core::WriteCellContext) -> Option<ExcelCellStyle> {
        // 自定义单元格样式
        None
    }
}

// 注册处理器
EasyExcel::write::<User>("output.xlsx")
    .register_write_handler(MyStyleHandler)
    .sheet("Sheet1")
    .do_write(data)?;
```

---

## 自定义转换器

```rust
use easyexcel_core::{Converter, ReadConverterContext, WriteConverterContext, CellValue, ExcelError};

struct YesNoConverter;

impl Converter<String> for YesNoConverter {
    fn support_excel_type(&self) -> easyexcel_core::CellDataType { easyexcel_core::CellDataType::String }
    
    fn convert_to_rust_data(&self, ctx: &ReadConverterContext) -> Result<String, ExcelError> {
        match ctx.raw_value() {
            CellValue::String(s) if s == "是" => Ok("YES".into()),
            CellValue::String(s) if s == "否" => Ok("NO".into()),
            other => Err(ExcelError::Format(format!("expected 是/否, got {other:?}")))
        }
    }

    fn convert_to_excel_data(&self, ctx: &WriteConverterContext<String>) -> Result<easyexcel_core::WriteCellData, ExcelError> {
        Ok(easyexcel_core::WriteCellData::from_string(
            if ctx.value() == "YES" { "是" } else { "否" }
        ))
    }
}
```

---

## 模块结构

| Crate | 功能 | Java 对应 |
|-------|------|-----------|
| `easyexcel` | 用户入口 Facade | `EasyExcel` / `EasyExcelFactory` |
| `easyexcel-core` | 核心 trait / 数据模型 / 错误类型 | `com.alibaba.excel.*` |
| `easyexcel-derive` | `#[derive(ExcelRow)]` 过程宏 | `@ExcelProperty` 注解处理 |
| `easyexcel-reader` | XLSX/XLS/CSV 读取引擎 | `analysis/` + `read/` |
| `easyexcel-writer` | XLSX/XLS/CSV 写入引擎 | `write/` |
| `easyexcel-template` | 模板填充引擎 | `write/metadata/fill/` |

---

## Java 兼容性

`easyexcel-rust` 与 Java EasyExcel 4.0.3 保持 1:1 对应：

- **335 个 Java @Test 方法** 全部有 Rust `#[test]` 对应
- **88 个 Golden 测试** 输出与 Java 完全一致
- **152 个 Parity 测试** 端到端行为等价
- 全量测试 **0 FAILEDs**

详见 [迁移文档](docs/migration/TEST_AUDIT_REPORT.md)。

---

## 许可证

Apache-2.0