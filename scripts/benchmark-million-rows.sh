#!/usr/bin/env sh
set -eu

rows="${1:-1000000}"
output="${2:-target/benchmark/million-rows.xlsx}"
cargo_command="${CARGO:-cargo}"

case "$(uname -s)" in
    Darwin)
        exec /usr/bin/time -l "$cargo_command" run --release -p easyexcel --example million_rows -- "$rows" "$output"
        ;;
    Linux)
        exec /usr/bin/time -v "$cargo_command" run --release -p easyexcel --example million_rows -- "$rows" "$output"
        ;;
    *)
        exec "$cargo_command" run --release -p easyexcel --example million_rows -- "$rows" "$output"
        ;;
esac
