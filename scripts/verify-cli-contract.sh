#!/usr/bin/env bash
# Verify that one built GridGrind artifact exposes the expected public help and protocol-catalog
# contract, including the interactive no-arg failure path. This is intentionally black-box: it
# only uses the artifact's own CLI surface.

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

resolve_cli_contract_python() {
    local candidates=()
    local candidate=''
    if [[ -x /usr/bin/python3 ]]; then
        candidates+=(/usr/bin/python3)
    fi
    if command -v python3 >/dev/null 2>&1; then
        candidate="$(command -v python3)"
        if [[ "${candidate}" != /usr/bin/python3 ]]; then
            candidates+=("${candidate}")
        fi
    fi

    for candidate in "${candidates[@]}"; do
        if "${candidate}" - <<'PY' >/dev/null 2>&1
import json
import pty
import select
PY
        then
            printf '%s\n' "${candidate}"
            return 0
        fi
    done

    die "python3 with json, pty, and select support is required for CLI contract verification"
}

readonly mode="${1:-}"
readonly target="${2:-}"
readonly script_dir="$(resolve_script_dir)"
readonly repo_root="$(cd -P -- "${script_dir}/.." && pwd)"
# shellcheck source=/dev/null
source "${repo_root}/scripts/lib/cli-contract-pty-support.sh"
readonly cli_contract_python="$(resolve_cli_contract_python)"
catalog_index_path=''
source_types_path=''
persistence_types_path=''
step_types_path=''
mutation_action_types_path=''
assertion_types_path=''
inspection_query_types_path=''
execution_mode_types_path=''
execution_policy_input_type_path=''
example_catalog_path=''
task_catalog_path=''
task_plan_path=''
task_keyword_match_report_path=''
doctor_report_path=''
request_template_path=''
help_overview_path=''
help_overview_stderr_path=''
help_protocol_path=''
help_protocol_stderr_path=''
help_guidance_path=''
help_guidance_stderr_path=''
noargs_stdout_path=''
noargs_stderr_path=''
temp_dir=''
temp_parent=''
doctor_execution_root=''

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
        interactive_launcher=("${target}")
        doctor_launcher=("${target}")
        doctor_execution_root=''
        label="binary ${target}"
        ;;
    jar)
        command -v java >/dev/null 2>&1 || die "java is required for jar verification"
        [[ -f "${target}" ]] || die "missing jar target: ${target}"
        launcher=(java -jar "${target}")
        interactive_launcher=(java -jar "${target}")
        doctor_launcher=(java -jar "${target}")
        doctor_execution_root=''
        label="jar ${target}"
        ;;
    docker-image)
        command -v docker >/dev/null 2>&1 || die "docker is required for docker-image verification"
        launcher=(docker run --rm "${target}")
        interactive_launcher=(docker run --rm -i -t "${target}")
        doctor_launcher=(docker run --rm -i "${target}")
        doctor_execution_root='/tmp'
        label="docker image ${target}"
        ;;
    *)
        die "unsupported mode ${mode}; expected binary, jar, or docker-image"
        ;;
esac

temp_parent="${repo_root}/tmp/verify-cli-contract"
mkdir -p "${temp_parent}"
temp_dir="${temp_parent}/run.$$.${RANDOM}"
rm -rf "${temp_dir}"
mkdir -p "${temp_dir}"
help_overview_path="${temp_dir}/help-overview.txt"
help_overview_stderr_path="${temp_dir}/help-overview.stderr"
help_protocol_path="${temp_dir}/help-protocol.txt"
help_protocol_stderr_path="${temp_dir}/help-protocol.stderr"
help_guidance_path="${temp_dir}/help-guidance.txt"
help_guidance_stderr_path="${temp_dir}/help-guidance.stderr"
noargs_stdout_path="${temp_dir}/noargs.stdout"
noargs_stderr_path="${temp_dir}/noargs.stderr"
catalog_index_path="${temp_dir}/protocol-catalog-index.json"
source_types_path="${temp_dir}/source-types.json"
persistence_types_path="${temp_dir}/persistence-types.json"
step_types_path="${temp_dir}/step-types.json"
mutation_action_types_path="${temp_dir}/mutation-action-types.json"
assertion_types_path="${temp_dir}/assertion-types.json"
inspection_query_types_path="${temp_dir}/inspection-query-types.json"
execution_mode_types_path="${temp_dir}/execution-mode-types.json"
execution_policy_input_type_path="${temp_dir}/execution-policy-input-type.json"
example_catalog_path="${temp_dir}/example-catalog.json"
task_catalog_path="${temp_dir}/task-catalog.json"
task_plan_path="${temp_dir}/task-plan.json"
task_keyword_match_report_path="${temp_dir}/task-keyword-match.json"
doctor_report_path="${temp_dir}/doctor-report.json"
request_template_path="${temp_dir}/request-template.json"
if [[ -z "${doctor_execution_root}" ]]; then
    doctor_execution_root="${temp_dir}"
fi
help_output="$("${launcher[@]}" --help 2> "${help_overview_stderr_path}" | tr -d '\r')"
help_stderr="$(tr -d '\r' < "${help_overview_stderr_path}")"
[[ -z "${help_stderr}" ]] || die "${label} --help wrote unexpected stderr: ${help_stderr}"
protocol_help_output="$("${launcher[@]}" --help-protocol 2> "${help_protocol_stderr_path}" | tr -d '\r')"
protocol_help_stderr="$(tr -d '\r' < "${help_protocol_stderr_path}")"
[[ -z "${protocol_help_stderr}" ]] || die "${label} --help-protocol wrote unexpected stderr: ${protocol_help_stderr}"
guidance_help_output="$("${launcher[@]}" --help-guidance 2> "${help_guidance_stderr_path}" | tr -d '\r')"
guidance_help_stderr="$(tr -d '\r' < "${help_guidance_stderr_path}")"
[[ -z "${guidance_help_stderr}" ]] || die "${label} --help-guidance wrote unexpected stderr: ${guidance_help_stderr}"
require_absent \
    "${protocol_help_output}" \
    'FORCE_FORMULA_RECALCULATION_ON_OPEN' \
    "${label} protocol help exposes the deleted recalc mutation action"
require_absent \
    "${protocol_help_output}" \
    'FORCE_FORMULA_RECALC_ON_OPEN' \
    "${label} protocol help exposes the rejected recalc shorthand"
require_contains \
    "${help_output}" \
    '--print-task-catalog' \
    "${label} overview help no longer advertises task-catalog printing"
require_contains \
    "${help_output}" \
    '--doctor-request' \
    "${label} overview help no longer advertises request doctoring"
require_contains \
    "${help_output}" \
    '--print-task-plan --lookup <id>' \
    "${label} overview help no longer advertises task-plan printing"
require_contains \
    "${help_output}" \
    '--print-task-keyword-match --query <text>' \
    "${label} overview help no longer advertises task-keyword-match printing"
require_contains \
    "${help_output}" \
    '--print-example --lookup <id>' \
    "${label} overview help no longer advertises built-in example printing"
require_contains \
    "${help_output}" \
    '--print-example-catalog' \
    "${label} overview help no longer advertises built-in example catalog printing"
require_contains \
    "${help_output}" \
    '--help-protocol' \
    "${label} overview help no longer advertises the protocol help surface"
require_contains \
    "${help_output}" \
    '--help-guidance' \
    "${label} overview help no longer advertises the guidance help surface"
require_contains \
    "${help_output}" \
    '--license' \
    "${label} overview help no longer advertises license rendering"
require_absent \
    "${help_output}" \
    'WARNING: A restricted method in java.lang.foreign.Linker has been called' \
    "${label} overview help leaked a Java native-access warning before product help"
require_absent \
    "${help_output}" \
    'Restricted methods will be blocked in a future release unless native access is enabled' \
    "${label} overview help leaked a Java native-access warning before product help"

printf '%s' "${help_output}" > "${help_overview_path}"
printf '%s' "${protocol_help_output}" > "${help_protocol_path}"
printf '%s' "${guidance_help_output}" > "${help_guidance_path}"

set +e
# Force the batch no-arg contract through a non-interactive stdin. The interactive PTY path is
# verified separately below, so inheriting a caller TTY here would make the verifier itself
# environment-sensitive instead of testing one stable surface.
"${launcher[@]}" < /dev/null > "${noargs_stdout_path}" 2> "${noargs_stderr_path}"
noargs_exit_code=$?
set -e
[[ ${noargs_exit_code} -eq 2 ]] || die "${label} bare invocation exited ${noargs_exit_code} instead of 2"
[[ ! -s "${noargs_stdout_path}" ]] || die "${label} bare invocation wrote unexpected stdout"
[[ -s "${noargs_stderr_path}" ]] || die "${label} bare invocation emitted no structured stderr"
verify_interactive_noarg_failure "${noargs_stderr_path}" "${interactive_launcher[@]}"

"${launcher[@]}" --print-request-template | tr -d '\r' > "${request_template_path}"
"${launcher[@]}" --print-protocol-catalog | tr -d '\r' > "${catalog_index_path}"
"${launcher[@]}" --print-protocol-catalog --lookup sourceTypes | tr -d '\r' > "${source_types_path}"
"${launcher[@]}" --print-protocol-catalog --lookup persistenceTypes | tr -d '\r' > "${persistence_types_path}"
"${launcher[@]}" --print-protocol-catalog --lookup stepTypes | tr -d '\r' > "${step_types_path}"
"${launcher[@]}" --print-protocol-catalog --lookup mutationActionTypes | tr -d '\r' > "${mutation_action_types_path}"
"${launcher[@]}" --print-protocol-catalog --lookup assertionTypes | tr -d '\r' > "${assertion_types_path}"
"${launcher[@]}" --print-protocol-catalog --lookup inspectionQueryTypes | tr -d '\r' > "${inspection_query_types_path}"
"${launcher[@]}" --print-protocol-catalog --lookup nestedTypes:executionModeTypes | tr -d '\r' > "${execution_mode_types_path}"
"${launcher[@]}" --print-protocol-catalog --lookup plainTypes:executionPolicyInputType | tr -d '\r' > "${execution_policy_input_type_path}"
"${launcher[@]}" --print-example-catalog | tr -d '\r' > "${example_catalog_path}"
"${launcher[@]}" --print-task-catalog | tr -d '\r' > "${task_catalog_path}"
"${launcher[@]}" --print-task-plan --lookup DASHBOARD | tr -d '\r' > "${task_plan_path}"
"${launcher[@]}" --print-task-keyword-match --query "monthly sales dashboard with charts" | tr -d '\r' > "${task_keyword_match_report_path}"
cat "${request_template_path}" \
    | "${doctor_launcher[@]}" --doctor-request --execution-root "${doctor_execution_root}" | tr -d '\r' > "${doctor_report_path}"

"${cli_contract_python}" - "${catalog_index_path}" "${source_types_path}" "${persistence_types_path}" "${step_types_path}" "${mutation_action_types_path}" "${assertion_types_path}" "${inspection_query_types_path}" "${execution_mode_types_path}" "${execution_policy_input_type_path}" "${example_catalog_path}" "${help_overview_path}" "${help_protocol_path}" "${help_guidance_path}" "${task_catalog_path}" "${task_plan_path}" "${task_keyword_match_report_path}" "${doctor_report_path}" "${request_template_path}" <<'PY'
import json
import sys
from pathlib import Path

catalog_index = json.loads(Path(sys.argv[1]).read_text())
source_types_group = json.loads(Path(sys.argv[2]).read_text())
persistence_types_group = json.loads(Path(sys.argv[3]).read_text())
step_types_group = json.loads(Path(sys.argv[4]).read_text())
mutation_action_types_group = json.loads(Path(sys.argv[5]).read_text())
assertion_types_group = json.loads(Path(sys.argv[6]).read_text())
inspection_query_types_group = json.loads(Path(sys.argv[7]).read_text())
execution_mode_types_group = json.loads(Path(sys.argv[8]).read_text())
execution_policy_input_type_group = json.loads(Path(sys.argv[9]).read_text())
example_catalog = json.loads(Path(sys.argv[10]).read_text())
overview_help_output = Path(sys.argv[11]).read_text()
protocol_help_output = Path(sys.argv[12]).read_text()
guidance_help_output = Path(sys.argv[13]).read_text()
task_catalog = json.loads(Path(sys.argv[14]).read_text())
task_plan = json.loads(Path(sys.argv[15]).read_text())
task_keyword_match_report = json.loads(Path(sys.argv[16]).read_text())
doctor_report = json.loads(Path(sys.argv[17]).read_text())
request_template = json.loads(Path(sys.argv[18]).read_text())

def die(message: str) -> None:
    print(f"error: {message}", file=sys.stderr)
    raise SystemExit(1)

def normalized_text(value: str) -> str:
    return " ".join(value.split())

top_level_groups = {entry["group"]: entry["entryIds"] for entry in catalog_index["topLevelGroups"]}
nested_type_groups = {entry["group"]: entry["entryIds"] for entry in catalog_index["nestedTypeGroups"]}
plain_type_groups = {entry["group"]: entry["entryIds"] for entry in catalog_index["plainTypeGroups"]}
lookup_namespaces = {entry["shape"]: entry["usage"] for entry in catalog_index["lookupNamespaces"]}
source_types = {entry["id"]: entry for entry in source_types_group["types"]}
persistence_types = {entry["id"]: entry for entry in persistence_types_group["types"]}
step_types = {entry["id"]: entry for entry in step_types_group["types"]}
mutation_action_types = {entry["id"]: entry for entry in mutation_action_types_group["types"]}
assertion_types = {entry["id"]: entry for entry in assertion_types_group["types"]}
inspection_query_types = {entry["id"]: entry for entry in inspection_query_types_group["types"]}
execution_mode_types = {
    entry["id"]: entry for entry in execution_mode_types_group["types"]
}
execution_policy_input_type = execution_policy_input_type_group["type"]
shipped_examples = example_catalog["examples"]
normalized_overview_help_output = normalized_text(overview_help_output)
normalized_protocol_help_output = normalized_text(protocol_help_output)
normalized_guidance_help_output = normalized_text(guidance_help_output)

required_protocol_help_snippets = (
    (
        "STREAMING_WRITE mode:",
        "protocol help no longer includes the CLI-owned STREAMING_WRITE limit label",
    ),
    (
        "Formula authoring:",
        "protocol help no longer includes the CLI-owned formula authoring limit label",
    ),
    (
        "array-formula braces such as {=SUM(A1:A2*B1:B2)} are rejected as INVALID_FORMULA",
        "protocol help no longer includes the CLI-owned formula authoring limit summary",
    ),
    (
        "ASSERTION steps for first-class verification",
        "protocol help no longer includes the CLI-owned step-kind summary",
    ),
    (
        "Step kind is inferred from exactly one of action, assertion, or query",
        "protocol help no longer explains that step kind is inferred without step.type",
    ),
)
required_guidance_help_snippets = (
    (
        "ANALYZE_WORKBOOK_FINDINGS aggregates",
        "guidance help no longer includes the CLI-owned workbook findings discovery line",
    ),
    (
        "gridgrind --print-example --lookup BUDGET --response example.json",
        "guidance help no longer includes the CLI-owned featured example command",
    ),
    (
        "requiredWorkspacePaths names those paths directly.",
        "guidance help no longer explains asset-backed built-in example portability",
    ),
)
for snippet, message in required_protocol_help_snippets:
    if normalized_text(snippet) not in normalized_protocol_help_output:
        die(message)
for snippet, message in required_guidance_help_snippets:
    if normalized_text(snippet) not in normalized_guidance_help_output:
        die(message)
if normalized_text("--print-task-catalog") not in normalized_overview_help_output:
    die("overview help no longer advertises task-catalog discovery")
if normalized_text("--doctor-request") not in normalized_overview_help_output:
    die("overview help no longer advertises request doctoring")
if normalized_text("--print-task-plan --lookup <id>") not in normalized_overview_help_output:
    die("overview help no longer advertises task-plan discovery")
if normalized_text("--print-task-keyword-match --query <text>") not in normalized_overview_help_output:
    die("overview help no longer advertises task-keyword-match discovery")
if normalized_text("--print-protocol-catalog --lookup <lookup-id>") not in normalized_overview_help_output:
    die("overview help no longer advertises the unified protocol-catalog lookup grammar")
if normalized_text("--print-example-catalog") not in normalized_overview_help_output:
    die("overview help no longer advertises example-catalog discovery")
if (
    normalized_text("--help-protocol") not in normalized_overview_help_output
    or normalized_text("--help-guidance") not in normalized_overview_help_output
):
    die("overview help no longer advertises the split help surfaces")
if normalized_text("--print-protocol-catalog --lookup <lookup-id>") not in normalized_protocol_help_output:
    die("protocol help no longer advertises the unified protocol-catalog lookup grammar")
if normalized_text("--print-protocol-catalog --lookup <id>|<group>:<id>") in normalized_overview_help_output:
    die("overview help still advertises the incomplete legacy protocol-catalog lookup grammar")
if normalized_text("--print-protocol-catalog --lookup <id>|<group>:<id>") in normalized_protocol_help_output:
    die("protocol help still advertises the incomplete legacy protocol-catalog lookup grammar")
if "--print-protocol-catalog --full" in normalized_overview_help_output:
    die("overview help still advertises the removed --full catalog surface")
if "--print-protocol-catalog --full" in normalized_protocol_help_output:
    die("protocol help still advertises the removed --full catalog surface")
if "-w /workdir" in guidance_help_output:
    die("guidance help still teaches the old Docker -w /workdir pattern")
if '-v "$(pwd)":/work' not in guidance_help_output:
    die("guidance help no longer teaches the mounted /work Docker pattern")

if catalog_index.get("protocolVersion") != "V1":
    die("protocol catalog index no longer emits protocolVersion=V1")
if catalog_index.get("discriminatorField") != "type":
    die("protocol catalog index no longer emits discriminatorField=type")
if catalog_index.get("requestTypeId") != "WorkbookPlan":
    die("protocol catalog index no longer emits requestTypeId=WorkbookPlan")
for required_group in ("mutationActionTypes", "assertionTypes", "inspectionQueryTypes"):
    if required_group not in top_level_groups:
        die(f"protocol catalog index no longer advertises top-level group {required_group}")
if "executionModeTypes" not in nested_type_groups:
    die("protocol catalog index no longer advertises nestedTypes:executionModeTypes")
if "executionPolicyInputType" not in plain_type_groups:
    die("protocol catalog index no longer advertises plainTypes:executionPolicyInputType")
if "<topLevelGroup>:<id>" not in lookup_namespaces:
    die("protocol catalog index no longer publishes the top-level lookup namespace")
if "nestedTypes:<group>" not in lookup_namespaces:
    die("protocol catalog index no longer publishes the nestedTypes lookup namespace")
if "plainTypes:<group>" not in lookup_namespaces:
    die("protocol catalog index no longer publishes the plainTypes lookup namespace")

if example_catalog.get("protocolVersion") != "V1":
    die("example catalog no longer emits protocolVersion=V1")
if not shipped_examples:
    die("example catalog examples is empty")
for example in shipped_examples:
    example_id = example["id"]
    request_file_name = example["requestFileName"]
    summary = example["summary"]
    guidance_snippet = normalized_text(
        f"- {example_id} requestFileName: {request_file_name} "
        + f"workspace: {example['workspaceMode']} summary: {summary}"
    )
    if guidance_snippet not in normalized_guidance_help_output:
        die(f"guidance help no longer lists the built-in example line for {example_id}")
    if set(example.keys()) != {"id", "requestFileName", "summary", "workspaceMode", "requiredWorkspacePaths"}:
        die(
            f"example catalog entry {example_id} exposes unexpected public fields: "
            + f"{sorted(example.keys())}"
        )
    if example["workspaceMode"] == "SELF_CONTAINED" and example["requiredWorkspacePaths"]:
        die(f"self-contained example {example_id} must not publish requiredWorkspacePaths")
    if example["workspaceMode"] == "REQUIRES_EXAMPLE_ASSETS" and not example["requiredWorkspacePaths"]:
        die(f"asset-backed example {example_id} must publish requiredWorkspacePaths")

expected_required_paths = {
        "CUSTOM_XML": [
        "custom-xml-assets/custom-xml-mapping.xlsx",
        "custom-xml-assets/custom-xml-update.xml",
    ],
    "SOURCE_BACKED_INPUT": [
        "source-backed-input-assets/title.txt",
        "source-backed-input-assets/total-formula.txt",
        "source-backed-input-assets/payload.bin",
    ],
    "PACKAGE_SECURITY_INSPECTION": [
        "package-security-assets/gridgrind-package-security.xlsx",
    ],
}
for example_id, required_paths in expected_required_paths.items():
    entry = next(example for example in shipped_examples if example["id"] == example_id)
    if entry["requiredWorkspacePaths"] != required_paths:
        die(
            f"example catalog requiredWorkspacePaths drifted for {example_id}: "
            + f"{entry['requiredWorkspacePaths']}"
        )
    required_paths_snippet = normalized_text(
        f"requiredWorkspacePaths: {', '.join(required_paths)}"
    )
    if required_paths_snippet not in normalized_guidance_help_output:
        die(f"guidance help no longer lists requiredWorkspacePaths for {example_id}")

execution_policy_summary = execution_policy_input_type["summary"]
if "execution.journal" not in execution_policy_summary:
    die("catalog executionPolicyInputType summary no longer advertises execution.journal")
if "execution.calculation" not in execution_policy_summary:
    die("catalog executionPolicyInputType summary no longer advertises execution.calculation")

for required_mode in ("FULL_XSSF", "EVENT_READ", "STREAMING_WRITE"):
    if required_mode not in execution_mode_types:
        die(f"catalog executionModeTypes no longer publishes {required_mode}")
execution_summary = execution_mode_types["STREAMING_WRITE"]["summary"]
for needle in ("DO_NOT_CALCULATE", "markRecalculateOnOpen=true", "ENSURE_SHEET", "APPEND_ROW"):
    if needle not in execution_summary:
        die(f"catalog executionModeTypes STREAMING_WRITE summary is missing '{needle}'")
if "FORCE_FORMULA_RECALCULATION_ON_OPEN" in execution_summary:
    die("catalog executionModeTypes STREAMING_WRITE summary exposes the deleted recalc mutation action")
if "FORCE_FORMULA_RECALC_ON_OPEN" in execution_summary:
    die("catalog executionModeTypes STREAMING_WRITE summary exposes the rejected removed recalc shorthand")
for streaming_term in ("ENSURE_SHEET", "APPEND_ROW"):
    if streaming_term not in protocol_help_output:
        die(f"protocol help no longer communicates STREAMING_WRITE constraint '{streaming_term}'")

set_range_template = mutation_action_types["SET_RANGE"]["stepTemplate"]["template"]["action"]["rows"]
if set_range_template.get("type") != "TYPED":
    die("catalog SET_RANGE step template no longer defaults to the TYPED cell-grid wrapper")
if "cells" not in set_range_template:
    die("catalog SET_RANGE step template no longer uses cells for the typed grid payload")

append_row_template = mutation_action_types["APPEND_ROW"]["stepTemplate"]["template"]["action"]["values"]
if append_row_template.get("type") != "TYPED":
    die("catalog APPEND_ROW step template no longer defaults to the TYPED cell-row wrapper")
if "cells" not in append_row_template:
    die("catalog APPEND_ROW step template no longer uses cells for the typed row payload")

if request_template.get("protocolVersion") != "V1":
    die("request template no longer emits protocolVersion=V1")
if request_template.get("source", {}).get("type") != "NEW":
    die("request template no longer emits source.type=NEW")
if request_template.get("persistence", {}).get("type") != "NONE":
    die("request template no longer emits persistence.type=NONE")
if "execution" in request_template:
    die("request template no longer omits the default execution block")
if "formulaEnvironment" in request_template:
    die("request template no longer omits the default formulaEnvironment block")
if request_template.get("steps") != []:
    die("request template no longer emits an empty steps list")

sheet_layout = inspection_query_types["GET_SHEET_LAYOUT"]["summary"]
if "presentation" not in sheet_layout:
    die("catalog GET_SHEET_LAYOUT summary no longer advertises layout.presentation")

formula_surface = inspection_query_types["GET_FORMULA_SURFACE"]["summary"]
if "surface.totalFormulaCellCount" not in formula_surface:
    die("catalog GET_FORMULA_SURFACE summary no longer describes grouped formula output")

formula_health = inspection_query_types["ANALYZE_FORMULA_HEALTH"]["summary"]
if "analysis.checkedFormulaCellCount" not in formula_health:
    die("catalog ANALYZE_FORMULA_HEALTH summary no longer advertises checked-count output")

named_range_surface = inspection_query_types["GET_NAMED_RANGE_SURFACE"]["summary"]
if "surface.workbookScopedCount" not in named_range_surface:
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

set_table_targets = mutation_action_types["SET_TABLE"].get("targetSelectors")
if set_table_targets != [{"family": "TableSelector", "typeIds": ["TABLE_BY_NAME_ON_SHEET"]}]:
    die("catalog SET_TABLE no longer publishes the exact allowed target selector family")

catalog_groups = {
    source_types_group["group"]: set(source_types),
    persistence_types_group["group"]: set(persistence_types),
    step_types_group["group"]: set(step_types),
    mutation_action_types_group["group"]: set(mutation_action_types),
    assertion_types_group["group"]: set(assertion_types),
    inspection_query_types_group["group"]: set(inspection_query_types),
}

tasks = task_catalog.get("tasks", [])
if not tasks:
    die("task catalog is empty")
for task in tasks:
    task_id = task.get("id")
    if not task_id:
        die("task catalog contains a task with no id")
    workflow = task.get("workflow", {})
    phases = workflow.get("phases", [])
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

if "task" in task_plan or "requestTemplate" in task_plan or "authoringNotes" in task_plan:
    die("task plan reintroduced the old wrapper shape instead of one direct request document")
if task_plan.get("source", {}).get("type") != "NEW":
    die("task plan no longer defaults DASHBOARD to a NEW source")
if task_plan.get("persistence", {}).get("type") != "SAVE_AS":
    die("task plan no longer defaults DASHBOARD to SAVE_AS persistence")
if task_plan.get("persistence", {}).get("ifExists") != "REPLACE":
    die("task plan no longer defaults DASHBOARD to SAVE_AS.ifExists=REPLACE")
if not task_plan.get("persistence", {}).get("path", "").endswith(".xlsx"):
    die("task plan no longer emits a syntactically valid SAVE_AS .xlsx path")
if "dashboard" not in task_plan.get("persistence", {}).get("path", ""):
    die("task plan no longer keeps the requested task id visible in the starter output path")
task_plan_steps = task_plan.get("steps", [])
if not isinstance(task_plan_steps, list):
    die("task plan no longer emits steps as a JSON array")
if not task_plan_steps:
    die("task plan no longer emits executable starter steps")
first_step = task_plan_steps[0]
if "stepId" not in first_step or "target" not in first_step:
    die("task plan starter steps no longer publish stepId and target placeholders")
if sum(1 for key in ("action", "assertion", "query") if key in first_step) != 1:
    die("task plan starter steps must expose exactly one step body")

if task_keyword_match_report.get("query") != "monthly sales dashboard with charts":
    die("task keyword match report no longer preserves the requested query text")
if task_keyword_match_report.get("candidates", []) == []:
    die("task keyword match report no longer returns ranked candidates")
first_candidate = task_keyword_match_report["candidates"][0]
if first_candidate.get("taskId") != "DASHBOARD":
    die("task keyword match report no longer ranks DASHBOARD first for a charted dashboard query")
if "dashboard" not in first_candidate.get("matchedTerms", []):
    die("task keyword match report no longer reports dashboard as a matched term")
if "chart" not in first_candidate.get("matchedTerms", []):
    die("task keyword match report no longer reports chart as a matched term")
if "summary" not in first_candidate or "score" not in first_candidate:
    die("task keyword match report no longer publishes compact candidate summary fields")
if not first_candidate.get("matchSources"):
    die("task keyword match report no longer publishes ranked match-source hints")
if "task" in first_candidate or "starterTemplate" in first_candidate or "reasons" in first_candidate:
    die("task keyword match report reintroduced the bulky embedded task payload")

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
if doctor_report.get("problems") != []:
    die("doctor report must emit an empty problems list for the minimal request")
PY

printf 'Verified CLI contract surface: %s\n' "${label}"
