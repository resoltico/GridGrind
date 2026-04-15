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

    if ! grep -Fq "${needle}" <<<"${text}"; then
        die "${description}"
    fi
}

require_absent() {
    local text=$1
    local needle=$2
    local description=$3

    if grep -Fq "${needle}" <<<"${text}"; then
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

cleanup() {
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
require_contains \
    "${help_output}" \
    'FORCE_FORMULA_RECALCULATION_ON_OPEN' \
    "${label} help output is missing the canonical streaming-write recalc operation name"
require_absent \
    "${help_output}" \
    'FORCE_FORMULA_RECALC_ON_OPEN' \
    "${label} help output still exposes the rejected legacy recalc shorthand"
require_contains \
    "${help_output}" \
    'LAMBDA/LET are currently rejected as INVALID_FORMULA' \
    "${label} help output no longer states the current hard LAMBDA/LET parser boundary"
require_contains \
    "${help_output}" \
    'ANALYZE_WORKBOOK_FINDINGS aggregates formula health, data-validation health, conditional-formatting health, autofilter health, table health, pivot-table health, hyperlink health, and named-range health.' \
    "${label} help output no longer explains the aggregate workbook-health scope"

catalog_path="$(mktemp "${TMPDIR:-/tmp}/gridgrind-cli-contract.XXXXXX.json")"
"${launcher[@]}" --print-protocol-catalog | tr -d '\r' > "${catalog_path}"

python3 - "${catalog_path}" <<'PY'
import json
import sys
from pathlib import Path

catalog = json.loads(Path(sys.argv[1]).read_text())

def die(message: str) -> None:
    print(f"error: {message}", file=sys.stderr)
    raise SystemExit(1)

plain_types = {entry["group"]: entry["type"] for entry in catalog["plainTypes"]}
read_types = {entry["id"]: entry for entry in catalog["readTypes"]}

execution_summary = plain_types["executionModeInputType"]["summary"]
if "FORCE_FORMULA_RECALCULATION_ON_OPEN" not in execution_summary:
    die("catalog executionModeInputType summary is missing the canonical recalc operation name")
if "FORCE_FORMULA_RECALC_ON_OPEN" in execution_summary:
    die("catalog executionModeInputType summary still exposes the rejected legacy recalc shorthand")

sheet_layout = read_types["GET_SHEET_LAYOUT"]["summary"]
if "presentation" not in sheet_layout:
    die("catalog GET_SHEET_LAYOUT summary no longer advertises layout.presentation")

formula_surface = read_types["GET_FORMULA_SURFACE"]["summary"]
if "analysis.totalFormulaCellCount" not in formula_surface:
    die("catalog GET_FORMULA_SURFACE summary no longer describes grouped formula output")

formula_health = read_types["ANALYZE_FORMULA_HEALTH"]["summary"]
if "analysis.checkedFormulaCellCount" not in formula_health:
    die("catalog ANALYZE_FORMULA_HEALTH summary no longer advertises checked-count output")

named_range_surface = read_types["GET_NAMED_RANGE_SURFACE"]["summary"]
if "analysis.workbookScopedCount" not in named_range_surface:
    die("catalog GET_NAMED_RANGE_SURFACE summary no longer advertises scope/count output")

named_range_health = read_types["ANALYZE_NAMED_RANGE_HEALTH"]["summary"]
if "analysis.checkedNamedRangeCount" not in named_range_health:
    die("catalog ANALYZE_NAMED_RANGE_HEALTH summary no longer advertises checked-count output")

workbook_findings = read_types["ANALYZE_WORKBOOK_FINDINGS"]["summary"]
for needle in (
    "all analysis families",
    "pivot-table health",
    "hyperlink health",
    "named-range health",
    "analysis.findings",
):
    if needle not in workbook_findings:
        die(f"catalog ANALYZE_WORKBOOK_FINDINGS summary is missing '{needle}'")
PY

printf 'Verified CLI contract surface: %s\n' "${label}"
