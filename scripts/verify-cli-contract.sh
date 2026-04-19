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
help_path=''

cleanup() {
    [[ -n "${help_path}" ]] && rm -f "${help_path}" || true
    [[ -n "${catalog_path}" ]] || return
    rm -f "${catalog_path}" || true
}

trap cleanup EXIT

[[ -n "${mode}" ]] || die "mode is required (binary, jar, or docker-image)"
[[ -n "${target}" ]] || die "target is required"

case "${mode}" in
    binary)
        [[ -x "${target}" ]] || die "binary target is not executable: ${target}"
        launcher=("${target}")
        label="binary ${target}"
        ;;
    jar)
        command -v java >/dev/null 2>&1 || die "java is required for jar verification"
        [[ -f "${target}" ]] || die "missing jar target: ${target}"
        launcher=(java -jar "${target}")
        label="jar ${target}"
        ;;
    docker-image)
        command -v docker >/dev/null 2>&1 || die "docker is required for docker-image verification"
        launcher=(docker run --rm "${target}")
        label="docker image ${target}"
        ;;
    *)
        die "unsupported mode ${mode}; expected binary, jar, or docker-image"
        ;;
esac

command -v python3 >/dev/null 2>&1 || die "python3 is required for CLI contract verification"

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
    '--print-example <id>' \
    "${label} help output no longer advertises built-in example printing"

help_path="$(mktemp "${TMPDIR:-/tmp}/gridgrind-cli-help.XXXXXX.txt")"
printf '%s\n' "${help_output}" > "${help_path}"

catalog_path="$(mktemp "${TMPDIR:-/tmp}/gridgrind-cli-contract.XXXXXX.json")"
"${launcher[@]}" --print-protocol-catalog | tr -d '\r' > "${catalog_path}"

python3 - "${catalog_path}" "${help_path}" <<'PY'
import json
import re
import sys
from pathlib import Path

catalog = json.loads(Path(sys.argv[1]).read_text())
help_output = Path(sys.argv[2]).read_text()

def die(message: str) -> None:
    print(f"error: {message}", file=sys.stderr)
    raise SystemExit(1)

plain_types = {entry["group"]: entry["type"] for entry in catalog["plainTypes"]}
inspection_query_types = {entry["id"]: entry for entry in catalog["inspectionQueryTypes"]}
assertion_types = {entry["id"]: entry for entry in catalog["assertionTypes"]}
cli_surface = catalog["cliSurface"]
shipped_examples = catalog["shippedExamples"]

for line in cli_surface["limitLines"]:
    if line.startswith("STREAMING_WRITE mode:") and line not in help_output:
        die("help output no longer includes the catalog-owned STREAMING_WRITE limit line")
    if line.startswith("Formula authoring:") and line not in help_output:
        die("help output no longer includes the catalog-owned formula authoring limit line")

for line in cli_surface["requestLines"]:
    if "ASSERTION steps for first-class verification" in line and line not in help_output:
        die("help output no longer includes the catalog-owned step-kind summary")

for line in cli_surface["discoveryLines"]:
    if "ANALYZE_WORKBOOK_FINDINGS" in line and line not in help_output:
        die("help output no longer includes the catalog-owned workbook findings discovery line")

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
PY

printf 'Verified CLI contract surface: %s\n' "${label}"
