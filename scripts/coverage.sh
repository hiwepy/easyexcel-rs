#!/usr/bin/env bash
set -euo pipefail

cargo llvm-cov clean --workspace
cargo llvm-cov \
  --workspace \
  --all-features \
  --ignore-filename-regex 'crates/easyexcel-derive/src/lib\.rs' \
  --fail-under-lines 100 \
  --fail-under-regions 100 \
  --fail-under-functions 100 \
  --html \
  --output-dir coverage
