#!/usr/bin/env bash
set -euo pipefail

cargo llvm-cov clean --workspace
cargo llvm-cov \
  --workspace \
  --all-features \
  --ignore-filename-regex 'crates/easyexcel-derive/src/lib\.rs' \
  --html \
  --output-dir coverage

# `cargo llvm-cov` does not enforce fail-under thresholds when it exports HTML.
# Its percentage threshold also rounds values such as 99.95% to 100, so require
# the raw missed counts in the TOTAL row to be exactly zero.
cargo llvm-cov report \
  --ignore-filename-regex 'crates/easyexcel-derive/src/lib\.rs' \
  --fail-under-lines 100 \
  --fail-under-regions 100 \
  --fail-under-functions 100 \
  --summary-only 2>&1 \
  | awk '{ print } $1 == "TOTAL" && ($3 != 0 || $6 != 0 || $9 != 0) { exit 1 }'
