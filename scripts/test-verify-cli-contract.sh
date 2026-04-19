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
  STREAMING_WRITE mode:     source.type must be NEW; mutation actions limited to ENSURE_SHEET and APPEND_ROW; execution.calculation must keep strategy=DO_NOT_CALCULATE and may set markRecalculateOnOpen=true.
  Formula authoring:        request-authored formulas are scalar only; array-formula braces such as {=SUM(A1:A2*B1:B2)} are rejected as INVALID_FORMULA, and LAMBDA/LET are currently rejected as INVALID_FORMULA because Apache POI cannot parse them. Other newer constructs may fail the same way.

Request:
  ANALYZE_WORKBOOK_FINDINGS aggregates formula health, data-validation health, conditional-formatting health, autofilter health, table health, pivot-table health, hyperlink health, and named-range health.
  Use MUTATION steps for workbook changes, ASSERTION steps for first-class verification, and INSPECTION steps for factual or analytical reads.

Discovery:
  Built-in generated examples:
    WORKBOOK_HEALTH  examples/workbook-health-request.json  Compact no-save workbook-health pass with targeted formula and aggregate findings.
    ASSERTION        examples/assertion-request.json  Mutate then verify with first-class assertions, verbose journaling, and factual readback.

Flags:
  --print-example <id>             Print one built-in generated example request.
'

readonly success_catalog='{
  "cliSurface": {
    "executionLines": [],
    "limitLines": [
      "STREAMING_WRITE mode:     source.type must be NEW; mutation actions limited to ENSURE_SHEET and APPEND_ROW; execution.calculation must keep strategy=DO_NOT_CALCULATE and may set markRecalculateOnOpen=true.",
      "Formula authoring:        request-authored formulas are scalar only; array-formula braces such as {=SUM(A1:A2*B1:B2)} are rejected as INVALID_FORMULA, and LAMBDA/LET are currently rejected as INVALID_FORMULA because Apache POI cannot parse them. Other newer constructs may fail the same way."
    ],
    "requestLines": [
      "Use MUTATION steps for workbook changes, ASSERTION steps for first-class verification, and INSPECTION steps for factual or analytical reads."
    ],
    "fileWorkflowLines": [],
    "coordinateSystems": [],
    "discoveryLines": [
      "ANALYZE_WORKBOOK_FINDINGS aggregates formula health, data-validation health, conditional-formatting health, autofilter health, table health, pivot-table health, hyperlink health, and named-range health."
    ],
    "standardInputRequiresRequestMessage": "STANDARD_INPUT-authored values require --request so stdin is available for input content instead of the request JSON"
  },
  "shippedExamples": [
    {
      "id": "WORKBOOK_HEALTH",
      "fileName": "workbook-health-request.json",
      "summary": "Compact no-save workbook-health pass with targeted formula and aggregate findings."
    },
    {
      "id": "ASSERTION",
      "fileName": "assertion-request.json",
      "summary": "Mutate then verify with first-class assertions, verbose journaling, and factual readback."
    }
  ],
  "plainTypes": [
    {
      "group": "executionPolicyInputType",
      "type": {
        "summary": "Optional request execution policy covering execution.mode, execution.journal, and execution.calculation. Omit it to accept the default FULL_XSSF read/write path, NORMAL journal detail, and DO_NOT_CALCULATE calculation policy."
      }
    },
    {
      "group": "executionModeInputType",
      "type": {
        "summary": "Execution-mode settings that select low-memory read and write execution families. readMode defaults to FULL_XSSF when omitted. writeMode defaults to FULL_XSSF when omitted. EVENT_READ supports GET_WORKBOOK_SUMMARY and GET_SHEET_SUMMARY only and requires execution.calculation.strategy=DO_NOT_CALCULATE with markRecalculateOnOpen=false (LIM-019). STREAMING_WRITE supports ENSURE_SHEET and APPEND_ROW on NEW workbooks only; execution.calculation may only keep strategy=DO_NOT_CALCULATE and optionally set markRecalculateOnOpen=true (LIM-020)."
      }
    }
  ],
  "assertionTypes": [
    {
      "id": "EXPECT_CELL_VALUE",
      "summary": "Require every selected cell to match one exact effective value."
    },
    {
      "id": "ALL_OF",
      "summary": "Require every nested assertion to pass against the same target."
    },
    {
      "id": "NOT",
      "summary": "Invert one nested assertion against the same target."
    }
  ],
  "inspectionQueryTypes": [
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
    "${success_help/markRecalculateOnOpen=true/FORCE_FORMULA_RECALC_ON_OPEN}" \
    "${success_catalog}"

run_verify_expect_failure \
    "${success_help/WORKBOOK_HEALTH/WORKBOOK_HEALTH_BROKEN}" \
    "${success_catalog}"

printf 'verify-cli-contract regression: success\n'
