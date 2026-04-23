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
    '')
        printf '%s' "${FAKE_GRIDGRIND_IMPLICIT_HELP:-${FAKE_GRIDGRIND_HELP:?}}"
        ;;
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

source "${repo_root}/scripts/lib/test-cli-contract-fixtures.sh"
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
    local implicit_help_text=${7:-${help_text}}
    if FAKE_GRIDGRIND_HELP="${help_text}" \
        FAKE_GRIDGRIND_IMPLICIT_HELP="${implicit_help_text}" \
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
    "${success_doctor_report/\"sourceType\": \"NEW\"/\"sourceType\": \"UTF8_FILE\"}" \
    "${success_help}"

run_verify_expect_failure \
    "${success_help}" \
    "${success_catalog}" \
    "${success_task_catalog}" \
    "${success_task_plan}" \
    "${success_goal_plan/\"id\": \"DASHBOARD\"/\"id\": \"TABULAR_REPORT\"}" \
    "${success_doctor_report}"

run_verify_expect_failure \
    "${success_help}" \
    "${success_catalog}" \
    "${success_task_catalog}" \
    "${success_task_plan}" \
    "${success_goal_plan}" \
    "${success_doctor_report}" \
    "${success_help/GridGrind 9.9.9/Implicit Help Drifted}"

printf 'verify-cli-contract regression: success\n'
