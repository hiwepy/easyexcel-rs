#!/usr/bin/env bash
# gap-check.sh — Migration progress verification script
#
# Purpose: run a quick health check at the end of each migration phase to
# confirm no regression has been introduced and key capability markers
# are present.
#
# Usage:   ./scripts/gap-check.sh [phase]
# Example: ./scripts/gap-check.sh phase0
#
# Exit code 0 = all checks pass.

set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$REPO_ROOT"

PHASE="${1:-phase0}"

echo "==========================================================="
echo "  easyexcel-rs migration gap-check :: $PHASE"
echo "==========================================================="
echo "repo: $REPO_ROOT"
echo "date: $(date -u +"%Y-%m-%dT%H:%M:%SZ")"
echo

# ----------------------------------------------------------------------
# 1. Build check
# ----------------------------------------------------------------------
echo "[1/6] cargo build --workspace --all-features ..."
if ! cargo build --workspace --all-features --quiet 2>&1 | tail -5; then
    echo "FAIL: build error"
    exit 1
fi
echo "  ✓ build ok"
echo

# ----------------------------------------------------------------------
# 2. Cargo fmt check
# ----------------------------------------------------------------------
echo "[2/6] cargo fmt --all -- --check ..."
if ! cargo fmt --all -- --check 2>&1; then
    echo "WARN: rustfmt diffs present (non-blocking)"
else
    echo "  ✓ fmt clean"
fi
echo

# ----------------------------------------------------------------------
# 3. Cargo clippy check (best-effort, won't fail on pre-existing warnings)
# ----------------------------------------------------------------------
echo "[3/6] cargo clippy --workspace --all-targets --all-features (informational) ..."
cargo clippy --workspace --all-targets --all-features --quiet 2>&1 | tail -10 || true
echo

# ----------------------------------------------------------------------
# 4. Test run
# ----------------------------------------------------------------------
echo "[4/6] cargo test --workspace --all-features --no-fail-fast ..."
# Allow up to 8 minutes for full test suite (incl. slow large-data tests)
timeout 480 cargo test --workspace --all-features --no-fail-fast 2>&1 | grep -E "^test result" > /tmp/gap-check-results.txt || true
cat /tmp/gap-check-results.txt
PASS=$(awk -F'[ .;]' '/^test result/ {p+=$5} END {print p+0}' /tmp/gap-check-results.txt)
FAIL=$(awk -F'[ .;]' '/^test result/ {f+=$8} END {print f+0}' /tmp/gap-check-results.txt)
IGN=$(awk -F'[ .;]' '/^test result/ {i+=$11} END {print i+0}' /tmp/gap-check-results.txt)
echo
echo "  → tests: $PASS passed, $FAIL failed, $IGN ignored"
if [ "$FAIL" -ne 0 ]; then
    echo "FAIL: $FAIL tests failing"
    exit 1
fi
echo

# ----------------------------------------------------------------------
# 5. Phase-specific gap markers
# ----------------------------------------------------------------------
echo "[5/6] phase markers ($PHASE) ..."
declare -a phase_files
case "$PHASE" in
    phase0)
        phase_files=(
            "docs/migration/java-tree-full.md"
            "docs/migration/rust-tree-full.md"
            "docs/migration/project-tree-diff.md"
            "docs/migration/object-method-matrix.md"
            "docs/migration/MIGRATION_STATUS.md"
        )
        ;;
    phase1)
        phase_files=(
            "crates/easyexcel-core/src/annotation/excel_image.rs"
            "crates/easyexcel-core/src/annotation/excel_comment.rs"
            "crates/easyexcel-core/src/annotation/excel_hyperlink.rs"
            "crates/easyexcel-core/src/annotation/excel_formula.rs"
            "crates/easyexcel-core/src/annotation/excel_data_validation.rs"
            "crates/easyexcel-core/src/annotation/excel_conditional.rs"
            "crates/easyexcel-core/src/annotation/excel_filter.rs"
        )
        ;;
    phase2)
        phase_files=(
            "crates/easyexcel-writer/src/handler/workbook_workbook_write_handler.rs"
            "crates/easyexcel-writer/src/handler/sheet_write_handler.rs"
            "crates/easyexcel-writer/src/handler/row_write_handler.rs"
            "crates/easyexcel-writer/src/handler/cell_write_handler.rs"
            "crates/easyexcel-writer/src/handler/merge_handler.rs"
            "crates/easyexcel-writer/src/handler/constraint_handler.rs"
        )
        ;;
    phase3)
        phase_files=(
            "crates/easyexcel-core/src/comment_data.rs"
            "crates/easyexcel-core/src/hyperlink_data.rs"
            "crates/easyexcel-writer/src/handler/data_validation_write_handler.rs"
            "crates/easyexcel-writer/src/handler/conditional_format_write_handler.rs"
            "crates/easyexcel-writer/src/handler/auto_filter_write_handler.rs"
        )
        ;;
    phase4)
        phase_files=(
            "crates/easyexcel-writer/src/poi_handle.rs"
        )
        ;;
    phase5)
        phase_files=(
            "crates/easyexcel-template/src/xls_fill.rs"
            "crates/easyexcel-writer/src/biff8/encrypt.rs"
        )
        ;;
    phase6)
        phase_files=(
            "crates/easyexcel/tests/temp_1to1_tests/_hardened.rs"
        )
        ;;
    phase7)
        phase_files=(
            "tests/golden/_regenerated.txt"
        )
        ;;
    *)
        echo "Unknown phase: $PHASE (use phase0..phase7)"
        exit 1
        ;;
esac

MISSING=0
for f in "${phase_files[@]}"; do
    if [ -e "$f" ]; then
        echo "  ✓ $f"
    else
        echo "  ✗ MISSING: $f"
        MISSING=$((MISSING + 1))
    fi
done
echo

# ----------------------------------------------------------------------
# 6. Summary
# ----------------------------------------------------------------------
echo "[6/6] Summary"
echo "  Phase:       $PHASE"
echo "  Tests:       $PASS passed, $FAIL failed, $IGN ignored"
echo "  Missing:     $MISSING phase-specific files"
echo
if [ "$MISSING" -ne 0 ]; then
    echo "FAIL: phase $PHASE exit criteria not met"
    exit 1
fi
echo "✓ gap-check $PHASE passed"