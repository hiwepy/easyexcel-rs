# Streaming benchmark baseline

The benchmark uses the public `EasyExcel` facade for both directions. It writes
typed rows with `constant_memory(true)`, then reads them through a listener that
counts and releases each row. The executable fails if the read count differs
from the requested row count.

## 2026-07-17 baseline

- Machine: Apple M4 Pro, 24 GiB RAM
- Operating system: macOS 26.5.2, arm64
- Rust: 1.93.1
- Profile: `release`
- Rows: 1,000,000 data rows plus one header row
- XLSX size: 12,336,908 bytes
- Write time: 2.927 seconds
- Read time: 0.647 seconds
- Whole command wall time: 4.31 seconds
- `/usr/bin/time -l` maximum resident set size: 8,519,680 bytes
- `/usr/bin/time -l` peak memory footprint: 2,408,856 bytes

The command was run after the release profile had been compiled, so whole
command wall time does not include a clean dependency build. Timing and memory
numbers are machine-specific; row-count verification is deterministic.

## Reproduce

```shell
./scripts/benchmark-million-rows.sh
```

The script uses `/usr/bin/time -l` on macOS and `/usr/bin/time -v` on Linux. Its
first argument overrides the row count and its second argument overrides the
output path.
