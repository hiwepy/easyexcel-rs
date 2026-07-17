//! Reproducible million-row constant-memory write and streaming-read benchmark.

use std::path::PathBuf;
use std::sync::Arc;
use std::sync::atomic::{AtomicUsize, Ordering};
use std::time::Instant;

use easyexcel::{
    AnalysisContext, CellValue, ConvertContext, EasyExcel, ExcelColumn, ExcelRow,
    ExcelWriteMetadata, FromExcelCell, IntoExcelCell, ReadListener, Result, RowData,
};

const DEFAULT_ROWS: u32 = 1_000_000;
const MAX_DATA_ROWS_WITH_ONE_HEADER: u32 = 1_048_575;

#[derive(Debug, ExcelRow)]
struct BenchmarkRow {
    #[excel(name = "ID", index = 0)]
    id: u32,
    #[excel(name = "Value", index = 1)]
    value: String,
}

struct CountingListener {
    rows: Arc<AtomicUsize>,
}

impl ReadListener<BenchmarkRow> for CountingListener {
    fn invoke(&mut self, data: BenchmarkRow, _context: &AnalysisContext) -> easyexcel::Result<()> {
        std::hint::black_box(data);
        self.rows.fetch_add(1, Ordering::Relaxed);
        Ok(())
    }
}

fn main() -> std::result::Result<(), Box<dyn std::error::Error>> {
    let rows = std::env::args()
        .nth(1)
        .map_or(Ok(DEFAULT_ROWS), |value| value.parse::<u32>())?;
    if rows > MAX_DATA_ROWS_WITH_ONE_HEADER {
        return Err(format!("row count {rows} exceeds the XLSX limit with one header row").into());
    }
    let path = std::env::args().nth(2).map_or_else(
        || PathBuf::from("target/benchmark/million-rows.xlsx"),
        PathBuf::from,
    );
    if let Some(parent) = path.parent() {
        std::fs::create_dir_all(parent)?;
    }

    let write_started = Instant::now();
    EasyExcel::write::<BenchmarkRow>(&path)
        .constant_memory(true)
        .do_write_iter((0..rows).map(|id| BenchmarkRow {
            id,
            value: format!("row-{id}"),
        }))?;
    let write_elapsed = write_started.elapsed();

    let observed = Arc::new(AtomicUsize::new(0));
    let read_started = Instant::now();
    EasyExcel::read::<BenchmarkRow, _>(
        &path,
        CountingListener {
            rows: Arc::clone(&observed),
        },
    )
    .do_read()?;
    let read_elapsed = read_started.elapsed();
    let observed = observed.load(Ordering::Relaxed);
    if observed != usize::try_from(rows)? {
        return Err(format!("expected {rows} rows, read {observed}").into());
    }

    println!("rows={rows}");
    println!("write_seconds={:.3}", write_elapsed.as_secs_f64());
    println!("read_seconds={:.3}", read_elapsed.as_secs_f64());
    println!("xlsx_bytes={}", std::fs::metadata(&path)?.len());
    println!("xlsx_path={}", path.display());
    Ok(())
}
