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

emit_fixture_file() {
    local fixture_path=$1
    [[ -f "${fixture_path}" ]] || {
        printf 'missing fixture file: %s\n' "${fixture_path}" >&2
        exit 1
    }
    cat "${fixture_path}"
}

case "${1:-}" in
    '')
        emit_fixture_file "${FAKE_GRIDGRIND_IMPLICIT_HELP_FILE:-${FAKE_GRIDGRIND_HELP_FILE:?}}"
        ;;
    --help)
        emit_fixture_file "${FAKE_GRIDGRIND_HELP_FILE:?}"
        ;;
    --doctor-request)
        emit_fixture_file "${FAKE_GRIDGRIND_DOCTOR_REPORT_FILE:?}"
        ;;
    --print-goal-plan)
        emit_fixture_file "${FAKE_GRIDGRIND_GOAL_PLAN_FILE:?}"
        ;;
    --print-task-catalog)
        emit_fixture_file "${FAKE_GRIDGRIND_TASK_CATALOG_FILE:?}"
        ;;
    --print-task-plan)
        emit_fixture_file "${FAKE_GRIDGRIND_TASK_PLAN_FILE:?}"
        ;;
    --print-protocol-catalog)
        emit_fixture_file "${FAKE_GRIDGRIND_CATALOG_FILE:?}"
        ;;
    *)
        printf 'unexpected invocation: %s\n' "$*" >&2
        exit 1
        ;;
esac
EOF
chmod +x "${fake_cli}"

source "${repo_root}/scripts/lib/test-cli-contract-fixtures.sh"
load_test_cli_contract_fixtures

verify_case_counter=0

next_case_dir() {
    verify_case_counter=$((verify_case_counter + 1))
    local case_dir="${test_root}/verify-case-${verify_case_counter}"
    mkdir -p "${case_dir}"
    printf '%s' "${case_dir}"
}

write_case_fixture() {
    local case_dir=$1
    local fixture_name=$2
    local fixture_text=$3
    local fixture_path="${case_dir}/${fixture_name}"

    printf '%s' "${fixture_text}" > "${fixture_path}"
    printf '%s' "${fixture_path}"
}

run_verify_with_fixture_texts() {
    local unset_tmpdir=$1
    local help_text=$2
    local catalog_text=$3
    local task_catalog_text=$4
    local task_plan_text=$5
    local goal_plan_text=$6
    local doctor_report_text=$7
    local implicit_help_text=$8
    local case_dir
    local help_file
    local catalog_file
    local task_catalog_file
    local task_plan_file
    local goal_plan_file
    local doctor_report_file
    local implicit_help_file

    case_dir="$(next_case_dir)"
    help_file="$(write_case_fixture "${case_dir}" 'help.txt' "${help_text}")"
    catalog_file="$(write_case_fixture "${case_dir}" 'protocol-catalog.json' "${catalog_text}")"
    task_catalog_file="$(write_case_fixture "${case_dir}" 'task-catalog.json' "${task_catalog_text}")"
    task_plan_file="$(write_case_fixture "${case_dir}" 'task-plan.json' "${task_plan_text}")"
    goal_plan_file="$(write_case_fixture "${case_dir}" 'goal-plan.json' "${goal_plan_text}")"
    doctor_report_file="$(write_case_fixture "${case_dir}" 'doctor-report.json' "${doctor_report_text}")"
    implicit_help_file="$(write_case_fixture "${case_dir}" 'implicit-help.txt' "${implicit_help_text}")"

    if [[ "${unset_tmpdir}" == 'true' ]]; then
        env -u TMPDIR \
            FAKE_GRIDGRIND_HELP_FILE="${help_file}" \
            FAKE_GRIDGRIND_IMPLICIT_HELP_FILE="${implicit_help_file}" \
            FAKE_GRIDGRIND_CATALOG_FILE="${catalog_file}" \
            FAKE_GRIDGRIND_TASK_CATALOG_FILE="${task_catalog_file}" \
            FAKE_GRIDGRIND_TASK_PLAN_FILE="${task_plan_file}" \
            FAKE_GRIDGRIND_GOAL_PLAN_FILE="${goal_plan_file}" \
            FAKE_GRIDGRIND_DOCTOR_REPORT_FILE="${doctor_report_file}" \
            "${verify_script}" binary "${fake_cli}" >/dev/null
        return 0
    fi

    FAKE_GRIDGRIND_HELP_FILE="${help_file}" \
        FAKE_GRIDGRIND_IMPLICIT_HELP_FILE="${implicit_help_file}" \
        FAKE_GRIDGRIND_CATALOG_FILE="${catalog_file}" \
        FAKE_GRIDGRIND_TASK_CATALOG_FILE="${task_catalog_file}" \
        FAKE_GRIDGRIND_TASK_PLAN_FILE="${task_plan_file}" \
        FAKE_GRIDGRIND_GOAL_PLAN_FILE="${goal_plan_file}" \
        FAKE_GRIDGRIND_DOCTOR_REPORT_FILE="${doctor_report_file}" \
        "${verify_script}" binary "${fake_cli}" >/dev/null
}

run_verify_expect_success() {
    run_verify_with_fixture_texts \
        false \
        "${success_help}" \
        "${success_catalog}" \
        "${success_task_catalog}" \
        "${success_task_plan}" \
        "${success_goal_plan}" \
        "${success_doctor_report}" \
        "${success_help}"
}

run_verify_expect_success_without_tmpdir() {
    run_verify_with_fixture_texts \
        true \
        "${success_help}" \
        "${success_catalog}" \
        "${success_task_catalog}" \
        "${success_task_plan}" \
        "${success_goal_plan}" \
        "${success_doctor_report}" \
        "${success_help}"
}

run_verify_expect_failure() {
    local help_text=$1
    local catalog_text=$2
    local task_catalog_text=${3:-${success_task_catalog}}
    local task_plan_text=${4:-${success_task_plan}}
    local goal_plan_text=${5:-${success_goal_plan}}
    local doctor_report_text=${6:-${success_doctor_report}}
    local implicit_help_text=${7:-${help_text}}
    if run_verify_with_fixture_texts \
        false \
        "${help_text}" \
        "${catalog_text}" \
        "${task_catalog_text}" \
        "${task_plan_text}" \
        "${goal_plan_text}" \
        "${doctor_report_text}" \
        "${implicit_help_text}" >/dev/null 2>&1; then
        die "verifier unexpectedly succeeded"
    fi
}

run_verify_expect_success
run_verify_expect_success_without_tmpdir

run_verify_expect_failure \
    "$(append_fixture_line "${success_help}" 'FORCE_FORMULA_RECALC_ON_OPEN')" \
    "${success_catalog}"

run_verify_expect_failure \
    "$(replace_fixture_token "${success_help}" 'WORKBOOK_HEALTH' 'WORKBOOK_HEALTH_BROKEN')" \
    "${success_catalog}"

run_verify_expect_failure \
    "${success_help}" \
    "${success_catalog}" \
    "$(replace_fixture_token "${success_task_catalog}" 'SET_TABLE' 'NO_SUCH_MUTATION')"

run_verify_expect_failure \
    "${success_help}" \
    "${success_catalog}" \
    "${success_task_catalog}" \
    "$(replace_fixture_token \
        "${success_task_plan}" \
        'dashboard-output.xlsx' \
        'dashboard-output.xls')"

run_verify_expect_failure \
    "${success_help}" \
    "${success_catalog}" \
    "${success_task_catalog}" \
    "${success_task_plan}" \
    "${success_goal_plan}" \
    "$(replace_fixture_token \
        "${success_doctor_report}" \
        '"sourceType" : "NEW"' \
        '"sourceType" : "UTF8_FILE"')" \
    "${success_help}"

run_verify_expect_failure \
    "${success_help}" \
    "${success_catalog}" \
    "${success_task_catalog}" \
    "${success_task_plan}" \
    "$(replace_fixture_token \
        "${success_goal_plan}" \
        '"id" : "DASHBOARD"' \
        '"id" : "TABULAR_REPORT"')" \
    "${success_doctor_report}"

run_verify_expect_failure \
    "${success_help}" \
    "${success_catalog}" \
    "${success_task_catalog}" \
    "${success_task_plan}" \
    "${success_goal_plan}" \
    "${success_doctor_report}" \
    "$(append_fixture_line "${success_help}" 'Implicit Help Drifted')"

printf 'verify-cli-contract regression: success\n'
