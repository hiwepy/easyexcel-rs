#!/usr/bin/env bash
# Export Java EasyExcel golden JSON expectations into crates/easyexcel/tests/golden/.
#
# Uses scripts/java-golden-export (Maven) against Alibaba EasyExcel 4.0.3.
# - Reads checked-in fixtures under crates/easyexcel/tests/fixtures
# - Writes SimpleData xlsx/csv artifacts under tests/golden/artifacts/
# - Emits *.expected.json (STRING-mode display cells) for Rust对照
#
# Dependencies:
#   - JDK 8+ (JAVA_HOME or Homebrew OpenJDK; EASYEXCEL_JAVA_HOME overrides)
#   - Apache Maven 3.6+ (`mvn` on PATH)
#   - Network once to resolve Maven deps (com.alibaba:easyexcel:4.0.3)
#
# Usage:
#   ./scripts/export-java-golden.sh
#   EASYEXCEL_JAVA_HOME=/path/to/jdk ./scripts/export-java-golden.sh
#   FIXTURES_DIR=/path/to/fixtures OUT_DIR=/path/to/golden ./scripts/export-java-golden.sh
#
# After export, commit updated tests/golden/*.expected.json (and artifacts/) so
# `cargo test -p easyexcel --test java_golden_tests` passes without a local JDK.

set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
EXPORT_DIR="${ROOT}/scripts/java-golden-export"
FIXTURES_DIR="${FIXTURES_DIR:-${ROOT}/crates/easyexcel/tests/fixtures}"
OUT_DIR="${OUT_DIR:-${ROOT}/crates/easyexcel/tests/golden}"

# Prefer Homebrew OpenJDK when plain `java` is missing from PATH (macOS).
if [[ -z "${JAVA_HOME:-}" ]]; then
  if [[ -n "${EASYEXCEL_JAVA_HOME:-}" ]]; then
    export JAVA_HOME="${EASYEXCEL_JAVA_HOME}"
  elif [[ -d /opt/homebrew/Cellar/openjdk ]]; then
    # shellcheck disable=SC2012
    LATEST="$(ls -1d /opt/homebrew/Cellar/openjdk/*/libexec/openjdk.jdk/Contents/Home 2>/dev/null | tail -1 || true)"
    if [[ -n "${LATEST}" ]]; then
      export JAVA_HOME="${LATEST}"
    fi
  elif command -v /usr/libexec/java_home >/dev/null 2>&1; then
    export JAVA_HOME="$(/usr/libexec/java_home 2>/dev/null || true)"
  fi
fi
if [[ -n "${JAVA_HOME:-}" ]]; then
  export PATH="${JAVA_HOME}/bin:${PATH}"
fi

if ! command -v mvn >/dev/null 2>&1; then
  echo "error: mvn not found; install Maven (https://maven.apache.org/) to export Java goldens" >&2
  exit 1
fi
if ! command -v java >/dev/null 2>&1; then
  echo "error: java not found; set JAVA_HOME or EASYEXCEL_JAVA_HOME (JDK 8+)" >&2
  exit 1
fi

if [[ ! -d "${FIXTURES_DIR}" ]]; then
  echo "error: fixtures dir missing: ${FIXTURES_DIR}" >&2
  exit 1
fi

mkdir -p "${OUT_DIR}"

echo "==> Java golden export"
echo "    fixtures: ${FIXTURES_DIR}"
echo "    out:      ${OUT_DIR}"
echo "    java:     $(java -version 2>&1 | head -1)"
echo "    mvn:      $(mvn -version 2>&1 | head -1)"

(
  cd "${EXPORT_DIR}"
  mvn -q -DskipTests package exec:java \
    -Dexec.mainClass=com.alibaba.easyexcel.golden.JavaGoldenExporter \
    -Dexec.args="${FIXTURES_DIR} ${OUT_DIR}"
)

echo "==> Done. Golden JSON:"
ls -1 "${OUT_DIR}"/*.expected.json 2>/dev/null || true
if [[ -d "${OUT_DIR}/artifacts" ]]; then
  echo "==> Artifacts (Java-written):"
  ls -1 "${OUT_DIR}/artifacts"/ 2>/dev/null || true
fi
