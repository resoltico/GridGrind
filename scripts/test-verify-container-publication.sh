#!/usr/bin/env bash
# Exercise the public-container verifier against a fake Docker CLI so the release workflow
# contract is tested locally without requiring a real registry push.

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
readonly verify_script="${repo_root}/scripts/verify-container-publication.sh"
readonly expected_description="$(
    awk -F= '
        $1 == "gridgrindDescription" {
            sub(/^[^=]*=/, "", $0)
            print $0
            exit
        }
    ' "${repo_root}/gradle.properties"
)"

[[ -x "${verify_script}" ]] || die "missing executable verifier script at ${verify_script}"
[[ -n "${expected_description}" ]] || die "missing gridgrindDescription in gradle.properties"

readonly temp_parent="${repo_root}/tmp/test-verify-container-publication"
mkdir -p "${temp_parent}"
test_root="${temp_parent}/run.$$"
rm -rf "${test_root}"
mkdir -p "${test_root}"
cleanup() {
    rm -rf "${test_root}"
}
trap cleanup EXIT

readonly fake_bin="${test_root}/bin"
readonly fake_log="${test_root}/docker.log"
mkdir -p "${fake_bin}"

cat > "${fake_bin}/docker" <<'EOF'
#!/usr/bin/env bash
set -euo pipefail

log_path=${FAKE_DOCKER_LOG:?}
printf '%s\n' "$*" >> "${log_path}"

args=("$@")
offset=0
if [[ ${#args[@]} -ge 2 && "${args[0]}" == "--config" ]]; then
    offset=2
fi

if [[ ${#args[@]} -le ${offset} ]]; then
    printf 'unexpected docker invocation: %s\n' "$*" >&2
    exit 1
fi

command=${args[${offset}]}
case "${command}" in
    pull)
        image_ref=${args[$((offset + 1))]:-}
        [[ -n "${image_ref}" ]] || exit 1
        exit 0
        ;;
    run)
        run_index=$((offset + 1))
        while [[ ${#args[@]} -gt ${run_index} ]]; do
            case "${args[${run_index}]}" in
                --rm|-i)
                    ((run_index += 1))
                    ;;
                *)
                    break
                    ;;
            esac
        done
        image_ref=${args[${run_index}]:-}
        cli_flag=${args[$((run_index + 1))]:-}
        case "${cli_flag}" in
            --version)
                case "${image_ref}" in
                    *:latest)
                        printf '%s' "${FAKE_DOCKER_LATEST_VERSION_OUTPUT:?}"
                        ;;
                    *)
                        printf '%s' "${FAKE_DOCKER_VERSION_OUTPUT:?}"
                        ;;
                esac
                ;;
            --help)
                printf '%s' "${FAKE_DOCKER_HELP_OUTPUT:?}"
                ;;
            --print-task-catalog)
                printf '%s' "${FAKE_DOCKER_TASK_CATALOG_OUTPUT:?}"
                ;;
            --print-task-plan)
                printf '%s' "${FAKE_DOCKER_TASK_PLAN_OUTPUT:?}"
                ;;
            --print-goal-plan)
                printf '%s' "${FAKE_DOCKER_GOAL_PLAN_OUTPUT:?}"
                ;;
            --doctor-request)
                printf '%s' "${FAKE_DOCKER_DOCTOR_REPORT_OUTPUT:?}"
                ;;
            --print-protocol-catalog)
                printf '%s' "${FAKE_DOCKER_CATALOG_OUTPUT:?}"
                ;;
            *)
                printf 'unexpected docker run invocation: %s\n' "$*" >&2
                exit 1
                ;;
        esac
        ;;
    *)
        printf 'unexpected docker subcommand: %s\n' "${command}" >&2
        exit 1
        ;;
esac
EOF
chmod +x "${fake_bin}/docker"

run_verify_expect_success() {
    PATH="${fake_bin}:${PATH}" \
        FAKE_DOCKER_LOG="${fake_log}" \
        FAKE_DOCKER_VERSION_OUTPUT="$1" \
        FAKE_DOCKER_LATEST_VERSION_OUTPUT="$2" \
        FAKE_DOCKER_HELP_OUTPUT="$3" \
        FAKE_DOCKER_CATALOG_OUTPUT="$4" \
        FAKE_DOCKER_TASK_CATALOG_OUTPUT="$5" \
        FAKE_DOCKER_TASK_PLAN_OUTPUT="$6" \
        FAKE_DOCKER_GOAL_PLAN_OUTPUT="$7" \
        FAKE_DOCKER_DOCTOR_REPORT_OUTPUT="$8" \
        GRIDGRIND_PUBLICATION_VERIFY_RETRIES=1 \
        GRIDGRIND_PUBLICATION_VERIFY_DELAY_SECONDS=0 \
        "${verify_script}" "ghcr.io/example/gridgrind" "9.9.9" >/dev/null
}

run_verify_expect_failure() {
    if PATH="${fake_bin}:${PATH}" \
        FAKE_DOCKER_LOG="${fake_log}" \
        FAKE_DOCKER_VERSION_OUTPUT="$1" \
        FAKE_DOCKER_LATEST_VERSION_OUTPUT="$2" \
        FAKE_DOCKER_HELP_OUTPUT="$3" \
        FAKE_DOCKER_CATALOG_OUTPUT="$4" \
        FAKE_DOCKER_TASK_CATALOG_OUTPUT="$5" \
        FAKE_DOCKER_TASK_PLAN_OUTPUT="$6" \
        FAKE_DOCKER_GOAL_PLAN_OUTPUT="$7" \
        FAKE_DOCKER_DOCTOR_REPORT_OUTPUT="$8" \
        GRIDGRIND_PUBLICATION_VERIFY_RETRIES=1 \
        GRIDGRIND_PUBLICATION_VERIFY_DELAY_SECONDS=0 \
        "${verify_script}" "ghcr.io/example/gridgrind" "9.9.9" >/dev/null 2>&1; then
        die "verifier unexpectedly succeeded"
    fi
}

expected_header="$(printf 'GridGrind 9.9.9\n%s' "${expected_description}")"
success_help='GridGrind 9.9.9
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
success_catalog='{
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

success_task_catalog='{
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

success_task_plan='{
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
    "steps": [
      {
        "stepId": "author-chart",
        "target": { "type": "BY_NAME", "name": "Dashboard" },
        "action": { "type": "SET_CHART" }
      },
      {
        "stepId": "verify-chart",
        "target": { "type": "ALL_ON_SHEET", "sheetName": "Dashboard" },
        "query": { "type": "GET_CHARTS" }
      }
    ]
  },
  "authoringNotes": [
    "The starter plan is runnable and seeds one simple dashboard chart workflow.",
    "Use --print-protocol-catalog --operation mutationActionTypes:SET_CHART for exact chart request shapes, or --print-protocol-catalog --search chart when browsing nearby surfaces.",
    "Replace any TODO-style .xlsx placeholder path before execution."
  ]
}'

success_goal_plan='{
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
          "steps": [
            {
              "stepId": "author-chart",
              "target": { "type": "BY_NAME", "name": "Dashboard" },
              "action": { "type": "SET_CHART" }
            },
            {
              "stepId": "verify-chart",
              "target": { "type": "ALL_ON_SHEET", "sheetName": "Dashboard" },
              "query": { "type": "GET_CHARTS" }
            }
          ]
        },
        "authoringNotes": [
          "The starter plan is runnable and seeds one simple dashboard chart workflow.",
          "Use --print-protocol-catalog --operation mutationActionTypes:SET_CHART for exact chart request shapes, or --print-protocol-catalog --search chart when browsing nearby surfaces.",
          "Replace any TODO-style .xlsx placeholder path before execution."
        ]
      }
    }
  ]
}'

success_doctor_report='{
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

: > "${fake_log}"
run_verify_expect_success \
    "${expected_header}" \
    "${expected_header}" \
    "${success_help}" \
    "${success_catalog}" \
    "${success_task_catalog}" \
    "${success_task_plan}" \
    "${success_goal_plan}" \
    "${success_doctor_report}"
grep -Fq 'pull ghcr.io/example/gridgrind:9.9.9' "${fake_log}" || die "verifier did not pull the version tag"
grep -Fq 'pull ghcr.io/example/gridgrind:latest' "${fake_log}" || die "verifier did not pull the latest tag"
grep -Fq 'run --rm ghcr.io/example/gridgrind:9.9.9 --help' "${fake_log}" || die \
    "verifier did not inspect the version tag help surface"
grep -Fq 'run --rm ghcr.io/example/gridgrind:latest --print-protocol-catalog' "${fake_log}" || die \
    "verifier did not inspect the latest tag catalog surface"
grep -Fq 'run --rm ghcr.io/example/gridgrind:latest --print-task-catalog' "${fake_log}" || die \
    "verifier did not inspect the latest tag task-catalog surface"
grep -Fq 'run --rm ghcr.io/example/gridgrind:latest --print-task-plan DASHBOARD' "${fake_log}" || die \
    "verifier did not inspect the latest tag task-plan surface"
grep -Fq 'run --rm ghcr.io/example/gridgrind:latest --print-goal-plan monthly sales dashboard with charts' "${fake_log}" || die \
    "verifier did not inspect the latest tag goal-plan surface"
grep -Fq 'run --rm -i ghcr.io/example/gridgrind:latest --doctor-request' "${fake_log}" || die \
    "verifier did not inspect the latest tag doctor surface"

run_verify_expect_failure \
    "$(printf 'gridgrind 9.9.9')" \
    "$(printf 'gridgrind 9.9.9')" \
    "${success_help}" \
    "${success_catalog}" \
    "${success_task_catalog}" \
    "${success_task_plan}" \
    "${success_goal_plan}" \
    "${success_doctor_report}"
run_verify_expect_failure "$(printf 'GridGrind 9.9.9\nWrong description')" \
    "$(printf 'GridGrind 9.9.9\nWrong description')" \
    "${success_help}" \
    "${success_catalog}" \
    "${success_task_catalog}" \
    "${success_task_plan}" \
    "${success_goal_plan}" \
    "${success_doctor_report}"
run_verify_expect_failure \
    "${expected_header}" \
    "${expected_header}" \
    "${success_help/markRecalculateOnOpen=true/FORCE_FORMULA_RECALC_ON_OPEN}" \
    "${success_catalog}" \
    "${success_task_catalog}" \
    "${success_task_plan}" \
    "${success_goal_plan}" \
    "${success_doctor_report}"
run_verify_expect_failure \
    "${expected_header}" \
    "${expected_header}" \
    "${success_help}" \
    "${success_catalog}" \
    "${success_task_catalog}" \
    "${success_task_plan}" \
    "${success_goal_plan}" \
    "${success_doctor_report/\"valid\": true/\"valid\": false}"
run_verify_expect_failure \
    "${expected_header}" \
    "${expected_header}" \
    "${success_help}" \
    "${success_catalog}" \
    "${success_task_catalog}" \
    "${success_task_plan}" \
    "${success_goal_plan/\"id\": \"DASHBOARD\"/\"id\": \"TABULAR_REPORT\"}" \
    "${success_doctor_report/\"valid\": true/\"valid\": false}"

printf 'verify-container-publication regression: success\n'
