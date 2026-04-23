#!/usr/bin/env bash
# Verify that one built GridGrind artifact exposes the expected public help and protocol-catalog
# contract. This is intentionally black-box: it only uses the artifact's own CLI surface.

set -euo pipefail

die() {
    printf 'error: %s\n' "$1" >&2
    exit 1
}

require_contains() {
    local text=$1
    local needle=$2
    local description=$3

    if ! grep -Fq -- "${needle}" <<<"${text}"; then
        die "${description}"
    fi
}

require_absent() {
    local text=$1
    local needle=$2
    local description=$3

    if grep -Fq -- "${needle}" <<<"${text}"; then
        die "${description}"
    fi
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

readonly mode="${1:-}"
readonly target="${2:-}"
readonly script_dir="$(resolve_script_dir)"
readonly repo_root="$(cd -P -- "${script_dir}/.." && pwd)"
catalog_path=''
task_catalog_path=''
task_plan_path=''
goal_plan_path=''
doctor_report_path=''
help_path=''
temp_dir=''
temp_parent=''

cleanup() {
    [[ -n "${temp_dir}" ]] && rm -rf "${temp_dir}" || true
}

trap cleanup EXIT

[[ -n "${mode}" ]] || die "mode is required (binary, jar, or docker-image)"
[[ -n "${target}" ]] || die "target is required"

case "${mode}" in
    binary)
        [[ -x "${target}" ]] || die "binary target is not executable: ${target}"
        launcher=("${target}")
        doctor_launcher=("${target}")
        label="binary ${target}"
        ;;
    jar)
        command -v java >/dev/null 2>&1 || die "java is required for jar verification"
        [[ -f "${target}" ]] || die "missing jar target: ${target}"
        launcher=(java -jar "${target}")
        doctor_launcher=(java -jar "${target}")
        label="jar ${target}"
        ;;
    docker-image)
        command -v docker >/dev/null 2>&1 || die "docker is required for docker-image verification"
        launcher=(docker run --rm "${target}")
        doctor_launcher=(docker run --rm -i "${target}")
        label="docker image ${target}"
        ;;
    *)
        die "unsupported mode ${mode}; expected binary, jar, or docker-image"
        ;;
esac

command -v python3 >/dev/null 2>&1 || die "python3 is required for CLI contract verification"

temp_parent="${repo_root}/tmp/verify-cli-contract"
mkdir -p "${temp_parent}"
temp_dir="${temp_parent}/run.$$.${RANDOM}"
rm -rf "${temp_dir}"
mkdir -p "${temp_dir}"
help_path="${temp_dir}/help.txt"
catalog_path="${temp_dir}/protocol-catalog.json"
task_catalog_path="${temp_dir}/task-catalog.json"
task_plan_path="${temp_dir}/task-plan.json"
goal_plan_path="${temp_dir}/goal-plan.json"
doctor_report_path="${temp_dir}/doctor-report.json"

help_output="$("${launcher[@]}" --help | tr -d '\r')"
require_absent \
    "${help_output}" \
    'FORCE_FORMULA_RECALCULATION_ON_OPEN' \
    "${label} help output still exposes the deleted recalc mutation action"
require_absent \
    "${help_output}" \
    'FORCE_FORMULA_RECALC_ON_OPEN' \
    "${label} help output still exposes the rejected legacy recalc shorthand"
require_contains \
    "${help_output}" \
    '--print-task-catalog' \
    "${label} help output no longer advertises task-catalog printing"
require_contains \
    "${help_output}" \
    '--doctor-request' \
    "${label} help output no longer advertises request doctoring"
require_contains \
    "${help_output}" \
    '--print-task-plan <id>' \
    "${label} help output no longer advertises task-plan printing"
require_contains \
    "${help_output}" \
    '--print-goal-plan <goal>' \
    "${label} help output no longer advertises goal-plan printing"
require_contains \
    "${help_output}" \
    '--print-example <id>' \
    "${label} help output no longer advertises built-in example printing"

printf '%s\n' "${help_output}" > "${help_path}"

"${launcher[@]}" --print-protocol-catalog | tr -d '\r' > "${catalog_path}"
"${launcher[@]}" --print-task-catalog | tr -d '\r' > "${task_catalog_path}"
"${launcher[@]}" --print-task-plan DASHBOARD | tr -d '\r' > "${task_plan_path}"
"${launcher[@]}" --print-goal-plan "monthly sales dashboard with charts" | tr -d '\r' > "${goal_plan_path}"
printf '%s\n' '{"source":{"type":"NEW"},"steps":[]}' \
    | "${doctor_launcher[@]}" --doctor-request | tr -d '\r' > "${doctor_report_path}"

python3 - "${catalog_path}" "${help_path}" "${task_catalog_path}" "${task_plan_path}" "${goal_plan_path}" "${doctor_report_path}" <<'PY'
import json
import re
import sys
from pathlib import Path

catalog = json.loads(Path(sys.argv[1]).read_text())
help_output = Path(sys.argv[2]).read_text()
task_catalog = json.loads(Path(sys.argv[3]).read_text())
task_plan = json.loads(Path(sys.argv[4]).read_text())
goal_plan = json.loads(Path(sys.argv[5]).read_text())
doctor_report = json.loads(Path(sys.argv[6]).read_text())

def die(message: str) -> None:
    print(f"error: {message}", file=sys.stderr)
    raise SystemExit(1)

plain_types = {entry["group"]: entry["type"] for entry in catalog["plainTypes"]}
inspection_query_types = {entry["id"]: entry for entry in catalog["inspectionQueryTypes"]}
assertion_types = {entry["id"]: entry for entry in catalog["assertionTypes"]}
cli_surface = catalog["cliSurface"]
shipped_examples = catalog["shippedExamples"]

for entry in cli_surface["limits"]["entries"]:
    label = entry["label"]
    value = entry["value"]
    if label == "STREAMING_WRITE mode":
        if f"{label}:" not in help_output or value not in help_output:
            die("help output no longer includes the catalog-owned STREAMING_WRITE limit entry")
    if label == "Formula authoring":
        if f"{label}:" not in help_output or value not in help_output:
            die("help output no longer includes the catalog-owned formula authoring limit entry")

for line in cli_surface["request"]["lines"]:
    if "ASSERTION steps for first-class verification" in line and line not in help_output:
        die("help output no longer includes the catalog-owned step-kind summary")
    if "do not send step.type" in line and line not in help_output:
        die("help output no longer explains that step kind is inferred without step.type")

for line in cli_surface["discovery"]["lines"]:
    if "ANALYZE_WORKBOOK_FINDINGS" in line and line not in help_output:
        die("help output no longer includes the catalog-owned workbook findings discovery line")

featured_example_command = cli_surface["discovery"]["printOneExampleCommand"]
if featured_example_command not in help_output:
    die("help output no longer includes the catalog-owned featured example command")
if "--print-task-catalog" not in help_output:
    die("help output no longer advertises task-catalog discovery")
if "--doctor-request" not in help_output:
    die("help output no longer advertises request doctoring")
if "--print-task-plan <id>" not in help_output:
    die("help output no longer advertises task-plan discovery")
if "--print-goal-plan <goal>" not in help_output:
    die("help output no longer advertises goal-plan discovery")

if not shipped_examples:
    die("catalog shippedExamples is empty")
for example in shipped_examples:
    example_id = example["id"]
    file_name = example["fileName"]
    summary = example["summary"]
    pattern = re.compile(
        rf"^\s*{re.escape(example_id)}\s+examples/{re.escape(file_name)}\s+{re.escape(summary)}\s*$",
        re.MULTILINE,
    )
    if not pattern.search(help_output):
        die(f"help output no longer lists the built-in example line for {example_id}")

execution_policy_summary = plain_types["executionPolicyInputType"]["summary"]
if "execution.journal" not in execution_policy_summary:
    die("catalog executionPolicyInputType summary no longer advertises execution.journal")
if "execution.calculation" not in execution_policy_summary:
    die("catalog executionPolicyInputType summary no longer advertises execution.calculation")

execution_summary = plain_types["executionModeInputType"]["summary"]
for needle in ("DO_NOT_CALCULATE", "markRecalculateOnOpen=true", "ENSURE_SHEET", "APPEND_ROW"):
    if needle not in execution_summary:
        die(f"catalog executionModeInputType summary is missing '{needle}'")
if "FORCE_FORMULA_RECALCULATION_ON_OPEN" in execution_summary:
    die("catalog executionModeInputType summary still exposes the deleted recalc mutation action")
if "FORCE_FORMULA_RECALC_ON_OPEN" in execution_summary:
    die("catalog executionModeInputType summary still exposes the rejected legacy recalc shorthand")

sheet_layout = inspection_query_types["GET_SHEET_LAYOUT"]["summary"]
if "presentation" not in sheet_layout:
    die("catalog GET_SHEET_LAYOUT summary no longer advertises layout.presentation")

formula_surface = inspection_query_types["GET_FORMULA_SURFACE"]["summary"]
if "analysis.totalFormulaCellCount" not in formula_surface:
    die("catalog GET_FORMULA_SURFACE summary no longer describes grouped formula output")

formula_health = inspection_query_types["ANALYZE_FORMULA_HEALTH"]["summary"]
if "analysis.checkedFormulaCellCount" not in formula_health:
    die("catalog ANALYZE_FORMULA_HEALTH summary no longer advertises checked-count output")

named_range_surface = inspection_query_types["GET_NAMED_RANGE_SURFACE"]["summary"]
if "analysis.workbookScopedCount" not in named_range_surface:
    die("catalog GET_NAMED_RANGE_SURFACE summary no longer advertises scope/count output")

named_range_health = inspection_query_types["ANALYZE_NAMED_RANGE_HEALTH"]["summary"]
if "analysis.checkedNamedRangeCount" not in named_range_health:
    die("catalog ANALYZE_NAMED_RANGE_HEALTH summary no longer advertises checked-count output")

workbook_findings = inspection_query_types["ANALYZE_WORKBOOK_FINDINGS"]["summary"]
for needle in (
    "all analysis families",
    "pivot-table health",
    "hyperlink health",
    "named-range health",
    "analysis.findings",
):
    if needle not in workbook_findings:
        die(f"catalog ANALYZE_WORKBOOK_FINDINGS summary is missing '{needle}'")

if "EXPECT_CELL_VALUE" not in assertion_types:
    die("catalog assertionTypes no longer includes EXPECT_CELL_VALUE")
if "ALL_OF" not in assertion_types or "NOT" not in assertion_types:
    die("catalog assertionTypes no longer includes the assertion composition operators")
if assertion_types["EXPECT_ANALYSIS_FINDING_PRESENT"].get("targetSelectorRule") != \
        "Matches the nested analysis query's target selectors.":
    die("catalog EXPECT_ANALYSIS_FINDING_PRESENT no longer publishes its derived target-selector rule")

mutation_action_types = {entry["id"]: entry for entry in catalog["mutationActionTypes"]}
set_table_targets = mutation_action_types["SET_TABLE"].get("targetSelectors")
if set_table_targets != [{"family": "TableSelector", "typeIds": ["BY_NAME_ON_SHEET"]}]:
    die("catalog SET_TABLE no longer publishes the exact allowed target selector family")

catalog_groups = {}
for key, value in catalog.items():
    if isinstance(value, list) and value and all(isinstance(entry, dict) for entry in value):
        ids = {entry["id"] for entry in value if "id" in entry}
        if ids:
            catalog_groups[key] = ids

tasks = task_catalog.get("tasks", [])
if not tasks:
    die("task catalog is empty")
for task in tasks:
    task_id = task.get("id")
    if not task_id:
        die("task catalog contains a task with no id")
    phases = task.get("phases", [])
    if not phases:
        die(f"task catalog task {task_id} contains no phases")
    for phase in phases:
        capability_refs = phase.get("capabilityRefs", [])
        if not capability_refs:
            die(f"task catalog task {task_id} phase {phase.get('label')} contains no capabilityRefs")
        for capability_ref in capability_refs:
            group = capability_ref.get("group")
            capability_id = capability_ref.get("id")
            if group not in catalog_groups:
                die(f"task catalog task {task_id} references unknown protocol group {group}")
            if capability_id not in catalog_groups[group]:
                die(
                    f"task catalog task {task_id} references unknown protocol capability "
                    f"{group}:{capability_id}"
                )

if task_plan.get("task", {}).get("id") != "DASHBOARD":
    die("task plan no longer resolves the requested task id")
task_plan_request = task_plan.get("requestTemplate", {})
if task_plan_request.get("source", {}).get("type") != "NEW":
    die("task plan no longer defaults DASHBOARD to a NEW source")
if task_plan_request.get("persistence", {}).get("type") != "SAVE_AS":
    die("task plan no longer defaults DASHBOARD to SAVE_AS persistence")
if not task_plan_request.get("persistence", {}).get("path", "").endswith(".xlsx"):
    die("task plan no longer emits a syntactically valid SAVE_AS .xlsx path")
task_plan_steps = task_plan_request.get("steps", [])
if not isinstance(task_plan_steps, list) or not task_plan_steps:
    die("task plan no longer emits a runnable starter scaffold with steps")
if not all(isinstance(step.get("stepId"), str) and step["stepId"] for step in task_plan_steps):
    die("task plan contains a step without a stable stepId")
if not any(step.get("action", {}).get("type") == "SET_CHART" for step in task_plan_steps):
    die("task plan no longer seeds the expected DASHBOARD chart authoring step")
if not any(step.get("query", {}).get("type") == "GET_CHARTS" for step in task_plan_steps):
    die("task plan no longer seeds the expected DASHBOARD chart verification step")
authoring_notes = task_plan.get("authoringNotes", [])
if not authoring_notes:
    die("task plan no longer publishes authoring notes")
if not any("--print-protocol-catalog --operation mutationActionTypes:SET_CHART" in note for note in authoring_notes):
    die("task plan no longer points authors back to exact chart capability lookups")
if not any("--print-protocol-catalog --search chart" in note for note in authoring_notes):
    die("task plan no longer points authors back to broader chart discovery")

if goal_plan.get("goal") != "monthly sales dashboard with charts":
    die("goal plan no longer preserves the requested goal text")
if goal_plan.get("candidates", []) == []:
    die("goal plan no longer returns ranked candidates")
first_candidate = goal_plan["candidates"][0]
if first_candidate.get("task", {}).get("id") != "DASHBOARD":
    die("goal plan no longer ranks DASHBOARD first for a charted dashboard goal")
if "dashboard" not in first_candidate.get("matchedTerms", []):
    die("goal plan no longer reports dashboard as a matched term")
if "chart" not in first_candidate.get("matchedTerms", []):
    die("goal plan no longer reports chart as a matched term")
if first_candidate.get("starterTemplate", {}).get("task", {}).get("id") != "DASHBOARD":
    die("goal plan no longer embeds the matching starter template")

if doctor_report.get("valid") is not True:
    die("doctor report no longer marks the minimal request as valid")
if doctor_report.get("severity") != "INFO":
    die("doctor report no longer classifies the minimal request as INFO")
doctor_summary = doctor_report.get("summary", {})
if doctor_summary.get("sourceType") != "NEW":
    die("doctor report no longer identifies the minimal request source type")
if doctor_summary.get("persistenceType") != "NONE":
    die("doctor report no longer identifies the minimal request persistence type")
if doctor_summary.get("stepCount") != 0:
    die("doctor report no longer reports the minimal request step count")
if doctor_summary.get("requiresStandardInputBinding") is not False:
    die("doctor report no longer reports STANDARD_INPUT requirements correctly")
if doctor_report.get("warnings") != []:
    die("doctor report no longer emits an empty warnings list for the minimal request")
if doctor_report.get("problem") is not None:
    die("doctor report must omit the blocking problem for valid requests")
PY

printf 'Verified CLI contract surface: %s\n' "${label}"
