#!/usr/bin/env bash
# Exercise the CLI contract verifier against a fake executable so the artifact-surface gate stays
# regression-tested without building a real jar or container image.

set -euo pipefail

die() {
    printf 'error: %s\n' "$1" >&2
    exit 1
}

resolve_script_dir() {
    local source_path="${BASH_SOURCE[0]}"
    while [[ -h "${source_path}" ]]; do
        local source_dir
        source_dir="$(cd -P -- "$(dirname -- "${source_path}")" && pwd)"
        source_path="$(readlink "${source_path}")"
        if [[ "${source_path}" != /* ]]; then
            source_path="${source_dir}/${source_path}"
        fi
    done
    cd -P -- "$(dirname -- "${source_path}")" && pwd
}

readonly script_dir="$(resolve_script_dir)"
readonly repo_root="$(cd -P -- "${script_dir}/.." && pwd)"
readonly verify_script="${repo_root}/scripts/verify-cli-contract.sh"

[[ -x "${verify_script}" ]] || die "missing executable verifier script at ${verify_script}"

test_root="$(mktemp -d)"
cleanup() {
    rm -rf "${test_root}"
}
trap cleanup EXIT

readonly fake_cli="${test_root}/gridgrind"

cat > "${fake_cli}" <<'EOF'
#!/usr/bin/env bash
set -euo pipefail

case "${1:-}" in
    --help)
        printf '%s' "${FAKE_GRIDGRIND_HELP:?}"
        ;;
    --print-protocol-catalog)
        printf '%s' "${FAKE_GRIDGRIND_CATALOG:?}"
        ;;
    *)
        printf 'unexpected invocation: %s\n' "$*" >&2
        exit 1
        ;;
esac
EOF
chmod +x "${fake_cli}"

readonly success_help='GridGrind 9.9.9
Structured .xlsx workbook automation engine with an agent-friendly JSON protocol

Limits:
  STREAMING_WRITE mode:     source.type must be NEW; operations limited to ENSURE_SHEET, APPEND_ROW, and FORCE_FORMULA_RECALCULATION_ON_OPEN.
  Formula authoring:        request-authored formulas are scalar only; array-formula braces such as {=SUM(A1:A2*B1:B2)} are rejected as INVALID_FORMULA, and LAMBDA/LET are currently rejected as INVALID_FORMULA because Apache POI cannot parse them. Other newer constructs may fail the same way.

Discovery:
  ANALYZE_WORKBOOK_FINDINGS aggregates formula health, data-validation health, conditional-formatting health, autofilter health, table health, pivot-table health, hyperlink health, and named-range health.
'

readonly success_catalog='{
  "plainTypes": [
    {
      "group": "executionModeInputType",
      "type": {
        "summary": "Optional top-level request settings that select low-memory read and write execution families. readMode defaults to FULL_XSSF when omitted. writeMode defaults to FULL_XSSF when omitted. EVENT_READ supports GET_WORKBOOK_SUMMARY and GET_SHEET_SUMMARY only (LIM-019). STREAMING_WRITE supports ENSURE_SHEET, APPEND_ROW, and FORCE_FORMULA_RECALCULATION_ON_OPEN on NEW workbooks only (LIM-020)."
      }
    }
  ],
  "readTypes": [
    {
      "id": "GET_SHEET_LAYOUT",
      "summary": "Return one sheet'\''s layout object with pane, zoomPercent, presentation, and per-row or per-column metadata."
    },
    {
      "id": "GET_FORMULA_SURFACE",
      "summary": "Return analysis.totalFormulaCellCount plus per-sheet formula usage groups."
    },
    {
      "id": "ANALYZE_FORMULA_HEALTH",
      "summary": "Return analysis.checkedFormulaCellCount, a severity summary, and findings."
    },
    {
      "id": "GET_NAMED_RANGE_SURFACE",
      "summary": "Return analysis.workbookScopedCount, sheetScopedCount, rangeBackedCount, formulaBackedCount, and namedRanges."
    },
    {
      "id": "ANALYZE_NAMED_RANGE_HEALTH",
      "summary": "Return analysis.checkedNamedRangeCount, a severity summary, and findings."
    },
    {
      "id": "ANALYZE_WORKBOOK_FINDINGS",
      "summary": "Return analysis.summary plus one flat analysis.findings list after running all analysis families (formula health, data-validation health, conditional-formatting health, autofilter health, table health, pivot-table health, hyperlink health, and named-range health) across the entire workbook."
    }
  ]
}'

run_verify_expect_success() {
    FAKE_GRIDGRIND_HELP="${success_help}" \
        FAKE_GRIDGRIND_CATALOG="${success_catalog}" \
        "${verify_script}" binary "${fake_cli}" >/dev/null
}

run_verify_expect_failure() {
    local help_text=$1
    local catalog_text=$2
    if FAKE_GRIDGRIND_HELP="${help_text}" \
        FAKE_GRIDGRIND_CATALOG="${catalog_text}" \
        "${verify_script}" binary "${fake_cli}" >/dev/null 2>&1; then
        die "verifier unexpectedly succeeded"
    fi
}

run_verify_expect_success

run_verify_expect_failure \
    "${success_help/FORCE_FORMULA_RECALCULATION_ON_OPEN/FORCE_FORMULA_RECALC_ON_OPEN}" \
    "${success_catalog}"

run_verify_expect_failure \
    "${success_help}" \
    "${success_catalog/presentation,/row-height,}"

printf 'verify-cli-contract regression: success\n'
