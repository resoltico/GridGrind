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
    --doctor-request)
        printf '%s' "${FAKE_GRIDGRIND_DOCTOR_REPORT:?}"
        ;;
    --print-goal-plan)
        printf '%s' "${FAKE_GRIDGRIND_GOAL_PLAN:?}"
        ;;
    --print-task-catalog)
        printf '%s' "${FAKE_GRIDGRIND_TASK_CATALOG:?}"
        ;;
    --print-task-plan)
        printf '%s' "${FAKE_GRIDGRIND_TASK_PLAN:?}"
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
  Request JSON size:        request JSON must not exceed 16 MiB (16777216 bytes); use UTF8_FILE, FILE, or STANDARD_INPUT sources for large authored payloads.
  STREAMING_WRITE mode:     source.type must be NEW; mutation actions limited to ENSURE_SHEET and APPEND_ROW; execution.calculation must keep strategy=DO_NOT_CALCULATE and may set markRecalculateOnOpen=true.
  Formula authoring:        request-authored formulas are scalar only; array-formula braces such as {=SUM(A1:A2*B1:B2)} are rejected as INVALID_FORMULA, and LAMBDA/LET are currently rejected as INVALID_FORMULA because Apache POI cannot parse them. Other newer constructs may fail the same way.

Request:
  ANALYZE_WORKBOOK_FINDINGS aggregates formula health, data-validation health, conditional-formatting health, autofilter health, table health, pivot-table health, hyperlink health, and named-range health.
  Large authored payloads belong in UTF8_FILE, FILE, or STANDARD_INPUT sources; the request JSON itself is capped at 16 MiB.
  Use MUTATION steps for workbook changes, ASSERTION steps for first-class verification, and INSPECTION steps for factual or analytical reads. Step kind is inferred from exactly one of action, assertion, or query; do not send step.type.

Discovery:
  gridgrind --print-task-catalog
  gridgrind --print-task-plan <id>
  gridgrind --print-goal-plan "monthly sales dashboard with charts"
  gridgrind --doctor-request [--request <path>]
  ANALYZE_WORKBOOK_FINDINGS aggregates formula health, data-validation health, conditional-formatting health, autofilter health, table health, pivot-table health, hyperlink health, and named-range health.
  Built-in generated examples:
    WORKBOOK_HEALTH  examples/workbook-health-request.json  Compact no-save workbook-health pass with targeted formula and aggregate findings.
    ASSERTION        examples/assertion-request.json  Mutate then verify with first-class assertions, verbose journaling, and factual readback.
  Print one built-in example:
    gridgrind --print-example WORKBOOK_HEALTH

Flags:
  --print-task-catalog           Print the machine-readable task catalog of high-level office-work recipes.
  --print-task-plan <id>         Print a machine-readable starter request scaffold for one task id.
  --print-goal-plan <goal>       Print ranked contract-owned task matches plus starter scaffolds for one freeform goal.
  --doctor-request               Lint one request and emit a machine-readable doctor report without opening or mutating a workbook.
  --print-example <id>             Print one built-in generated example request.
'

readonly success_catalog='{
  "cliSurface": {
    "usage": { "label": "Usage", "lines": [] },
    "execution": { "label": "Execution", "lines": [] },
    "limits": {
      "label": "Limits",
      "entries": [
        {
          "label": "Request JSON size",
          "value": "request JSON must not exceed 16 MiB (16777216 bytes); use UTF8_FILE, FILE, or STANDARD_INPUT sources for large authored payloads."
        },
        {
          "label": "STREAMING_WRITE mode",
          "value": "source.type must be NEW; mutation actions limited to ENSURE_SHEET and APPEND_ROW; execution.calculation must keep strategy=DO_NOT_CALCULATE and may set markRecalculateOnOpen=true."
        },
        {
          "label": "Formula authoring",
          "value": "request-authored formulas are scalar only; array-formula braces such as {=SUM(A1:A2*B1:B2)} are rejected as INVALID_FORMULA, and LAMBDA/LET are currently rejected as INVALID_FORMULA because Apache POI cannot parse them. Other newer constructs may fail the same way."
        }
      ]
    },
    "request": {
      "label": "Request",
      "lines": [
        "Large authored payloads belong in UTF8_FILE, FILE, or STANDARD_INPUT sources; the request JSON itself is capped at 16 MiB.",
        "Use MUTATION steps for workbook changes, ASSERTION steps for first-class verification, and INSPECTION steps for factual or analytical reads. Step kind is inferred from exactly one of action, assertion, or query; do not send step.type."
      ]
    },
    "fileWorkflow": { "label": "File Workflow", "entries": [] },
    "coordinateSystems": {
      "label": "Coordinate Systems",
      "leftHeader": "Pattern",
      "rightHeader": "Convention / Example",
      "entries": []
    },
    "minimalValidRequest": { "label": "Minimal Valid Request" },
    "stdinExample": { "label": "stdin Example", "commandLines": [], "description": null },
    "dockerFileExample": { "label": "Docker File Example", "commandLines": [], "description": null },
    "discovery": {
      "label": "Discovery",
      "lines": [
        "gridgrind --print-task-catalog",
        "gridgrind --print-task-plan <id>",
        "ANALYZE_WORKBOOK_FINDINGS aggregates formula health, data-validation health, conditional-formatting health, autofilter health, table health, pivot-table health, hyperlink health, and named-range health."
      ],
      "builtInExamplesLabel": "Built-in generated examples",
      "printOneExampleLabel": "Print one built-in example",
      "protocolCatalogNote": "note",
      "printOneExampleCommand": "gridgrind --print-example WORKBOOK_HEALTH"
    },
    "docs": { "label": "Docs", "entries": [] },
    "flags": { "label": "Flags", "entries": [] },
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
  "mutationActionTypes": [
    {
      "id": "SET_TABLE",
      "summary": "Create or replace one workbook-global table definition.",
      "targetSelectors": [
        { "family": "TableSelector", "typeIds": ["BY_NAME_ON_SHEET"] }
      ]
    }
  ],
  "assertionTypes": [
    {
      "id": "EXPECT_CELL_VALUE",
      "summary": "Require every selected cell to match one exact effective value."
    },
    {
      "id": "EXPECT_ANALYSIS_FINDING_PRESENT",
      "summary": "Run one analysis query against the selected target and require one matching finding.",
      "targetSelectorRule": "Matches the nested analysis query'\''s target selectors."
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

readonly success_task_catalog='{
  "protocolVersion": "V1",
  "tasks": [
    {
      "id": "REPORT",
      "summary": "Build a report.",
      "intentTags": ["office"],
      "outcomes": ["report saved"],
      "requiredInputs": ["data"],
      "optionalFeatures": ["assertions"],
      "phases": [
        {
          "label": "Shape",
          "objective": "Lay out the report.",
          "capabilityRefs": [
            { "group": "mutationActionTypes", "id": "SET_TABLE" }
          ],
          "notes": ["note"]
        },
        {
          "label": "Verify",
          "objective": "Validate and inspect the report.",
          "capabilityRefs": [
            { "group": "assertionTypes", "id": "EXPECT_CELL_VALUE" },
            { "group": "inspectionQueryTypes", "id": "ANALYZE_WORKBOOK_FINDINGS" }
          ],
          "notes": ["note"]
        }
      ],
      "commonPitfalls": ["pitfall"]
    }
  ]
}'

readonly success_task_plan='{
  "protocolVersion": "V1",
  "task": {
    "id": "DASHBOARD",
    "summary": "Build a dashboard.",
    "intentTags": ["office"],
    "outcomes": ["dashboard saved"],
    "requiredInputs": ["metrics"],
    "optionalFeatures": ["assertions"],
    "phases": [
      {
        "label": "Author",
        "objective": "Create the dashboard.",
        "capabilityRefs": [
          { "group": "sourceTypes", "id": "NEW" },
          { "group": "persistenceTypes", "id": "SAVE_AS" },
          { "group": "mutationActionTypes", "id": "SET_TABLE" }
        ],
        "notes": ["note"]
      }
    ],
    "commonPitfalls": ["pitfall"]
  },
  "requestTemplate": {
    "source": { "type": "NEW" },
    "persistence": { "type": "SAVE_AS", "path": "todo-dashboard-output.xlsx" },
    "steps": []
  },
  "authoringNotes": [
    "requestTemplate is intentionally minimal and valid: source and persistence are scaffolded, but steps stays empty until you author the workflow.",
    "Use task.phases[*].capabilityRefs to discover the exact operation shapes through --print-protocol-catalog --operation <group>:<id>.",
    "Replace any TODO-style .xlsx placeholder path before execution."
  ]
}'

readonly success_goal_plan='{
  "protocolVersion": "V1",
  "goal": "monthly sales dashboard with charts",
  "normalizedTerms": ["monthly", "sale", "dashboard", "chart"],
  "unmatchedTerms": ["monthly", "sale"],
  "suggestedIntentTags": ["office", "dashboard", "charts", "summary"],
  "candidates": [
    {
      "task": {
        "id": "DASHBOARD",
        "summary": "Build a dashboard.",
        "intentTags": ["office", "dashboard", "charts", "summary"],
        "outcomes": ["dashboard saved"],
        "requiredInputs": ["metrics"],
        "optionalFeatures": ["assertions"],
        "phases": [
          {
            "label": "Author",
            "objective": "Create the dashboard.",
            "capabilityRefs": [
              { "group": "sourceTypes", "id": "NEW" },
              { "group": "persistenceTypes", "id": "SAVE_AS" },
              { "group": "mutationActionTypes", "id": "SET_CHART" }
            ],
            "notes": ["note"]
          }
        ],
        "commonPitfalls": ["pitfall"]
      },
      "score": 26,
      "matchedTerms": ["dashboard", "chart"],
      "reasons": [
        "Matched intent tag \"dashboard\" via dashboard.",
        "Matched capability \"set chart\" via chart."
      ],
      "starterTemplate": {
        "protocolVersion": "V1",
        "task": {
          "id": "DASHBOARD",
          "summary": "Build a dashboard.",
          "intentTags": ["office", "dashboard", "charts", "summary"],
          "outcomes": ["dashboard saved"],
          "requiredInputs": ["metrics"],
          "optionalFeatures": ["assertions"],
          "phases": [
            {
              "label": "Author",
              "objective": "Create the dashboard.",
              "capabilityRefs": [
                { "group": "sourceTypes", "id": "NEW" },
                { "group": "persistenceTypes", "id": "SAVE_AS" },
                { "group": "mutationActionTypes", "id": "SET_CHART" }
              ],
              "notes": ["note"]
            }
          ],
          "commonPitfalls": ["pitfall"]
        },
        "requestTemplate": {
          "source": { "type": "NEW" },
          "persistence": { "type": "SAVE_AS", "path": "todo-dashboard-output.xlsx" },
          "steps": []
        },
        "authoringNotes": [
          "requestTemplate is intentionally minimal and valid: source and persistence are scaffolded, but steps stays empty until you author the workflow.",
          "Use task.phases[*].capabilityRefs to discover the exact operation shapes through --print-protocol-catalog --operation <group>:<id>.",
          "Replace any TODO-style .xlsx placeholder path before execution."
        ]
      }
    }
  ]
}'

readonly success_doctor_report='{
  "protocolVersion": "V1",
  "severity": "INFO",
  "valid": true,
  "summary": {
    "sourceType": "NEW",
    "persistenceType": "NONE",
    "readMode": "FULL_XSSF",
    "writeMode": "FULL_XSSF",
    "calculationStrategy": "DO_NOT_CALCULATE",
    "markRecalculateOnOpen": false,
    "requiresStandardInputBinding": false,
    "stepCount": 0,
    "mutationStepCount": 0,
    "assertionStepCount": 0,
    "inspectionStepCount": 0
  },
  "warnings": [],
  "problem": null
}'

run_verify_expect_success() {
    FAKE_GRIDGRIND_HELP="${success_help}" \
        FAKE_GRIDGRIND_CATALOG="${success_catalog}" \
        FAKE_GRIDGRIND_TASK_CATALOG="${success_task_catalog}" \
        FAKE_GRIDGRIND_TASK_PLAN="${success_task_plan}" \
        FAKE_GRIDGRIND_GOAL_PLAN="${success_goal_plan}" \
        FAKE_GRIDGRIND_DOCTOR_REPORT="${success_doctor_report}" \
        "${verify_script}" binary "${fake_cli}" >/dev/null
}

run_verify_expect_success_without_tmpdir() {
    env -u TMPDIR \
        FAKE_GRIDGRIND_HELP="${success_help}" \
        FAKE_GRIDGRIND_CATALOG="${success_catalog}" \
        FAKE_GRIDGRIND_TASK_CATALOG="${success_task_catalog}" \
        FAKE_GRIDGRIND_TASK_PLAN="${success_task_plan}" \
        FAKE_GRIDGRIND_GOAL_PLAN="${success_goal_plan}" \
        FAKE_GRIDGRIND_DOCTOR_REPORT="${success_doctor_report}" \
        "${verify_script}" binary "${fake_cli}" >/dev/null
}

run_verify_expect_failure() {
    local help_text=$1
    local catalog_text=$2
    local task_catalog_text=${3:-${success_task_catalog}}
    local task_plan_text=${4:-${success_task_plan}}
    local goal_plan_text=${5:-${success_goal_plan}}
    local doctor_report_text=${6:-${success_doctor_report}}
    if FAKE_GRIDGRIND_HELP="${help_text}" \
        FAKE_GRIDGRIND_CATALOG="${catalog_text}" \
        FAKE_GRIDGRIND_TASK_CATALOG="${task_catalog_text}" \
        FAKE_GRIDGRIND_TASK_PLAN="${task_plan_text}" \
        FAKE_GRIDGRIND_GOAL_PLAN="${goal_plan_text}" \
        FAKE_GRIDGRIND_DOCTOR_REPORT="${doctor_report_text}" \
        "${verify_script}" binary "${fake_cli}" >/dev/null 2>&1; then
        die "verifier unexpectedly succeeded"
    fi
}

run_verify_expect_success
run_verify_expect_success_without_tmpdir

run_verify_expect_failure \
    "${success_help/markRecalculateOnOpen=true/FORCE_FORMULA_RECALC_ON_OPEN}" \
    "${success_catalog}"

run_verify_expect_failure \
    "${success_help/WORKBOOK_HEALTH/WORKBOOK_HEALTH_BROKEN}" \
    "${success_catalog}"

run_verify_expect_failure \
    "${success_help}" \
    "${success_catalog}" \
    "${success_task_catalog/SET_TABLE/NO_SUCH_MUTATION}"

run_verify_expect_failure \
    "${success_help}" \
    "${success_catalog}" \
    "${success_task_catalog}" \
    "${success_task_plan/todo-dashboard-output.xlsx/todo-dashboard-output.xls}"

run_verify_expect_failure \
    "${success_help}" \
    "${success_catalog}" \
    "${success_task_catalog}" \
    "${success_task_plan}" \
    "${success_goal_plan}" \
    "${success_doctor_report/\"sourceType\": \"NEW\"/\"sourceType\": \"UTF8_FILE\"}"

run_verify_expect_failure \
    "${success_help}" \
    "${success_catalog}" \
    "${success_task_catalog}" \
    "${success_task_plan}" \
    "${success_goal_plan/\"id\": \"DASHBOARD\"/\"id\": \"TABULAR_REPORT\"}" \
    "${success_doctor_report}"

printf 'verify-cli-contract regression: success\n'
