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
        if [[ -t 0 || -t 1 || -t 2 ]]; then
            emit_fixture_file "${FAKE_GRIDGRIND_INTERACTIVE_NOARGS_FAILURE_FILE:-${FAKE_GRIDGRIND_NOARGS_FAILURE_FILE:?}}"
        else
            emit_fixture_file "${FAKE_GRIDGRIND_NOARGS_FAILURE_FILE:?}" >&2
        fi
        exit 2
        ;;
    --help)
        emit_fixture_file "${FAKE_GRIDGRIND_HELP_OVERVIEW_FILE:?}"
        ;;
    --help-protocol)
        emit_fixture_file "${FAKE_GRIDGRIND_HELP_PROTOCOL_FILE:?}"
        ;;
    --help-guidance)
        emit_fixture_file "${FAKE_GRIDGRIND_HELP_GUIDANCE_FILE:?}"
        ;;
    --print-request-template)
        emit_fixture_file "${FAKE_GRIDGRIND_REQUEST_TEMPLATE_FILE:?}"
        ;;
    --doctor-request)
        emit_fixture_file "${FAKE_GRIDGRIND_DOCTOR_REPORT_FILE:?}"
        ;;
    --print-task-keyword-match)
        emit_fixture_file "${FAKE_GRIDGRIND_TASK_KEYWORD_MATCH_FILE:?}"
        ;;
    --print-task-catalog)
        emit_fixture_file "${FAKE_GRIDGRIND_TASK_CATALOG_FILE:?}"
        ;;
    --print-example-catalog)
        emit_fixture_file "${FAKE_GRIDGRIND_EXAMPLE_CATALOG_FILE:?}"
        ;;
    --print-task-plan)
        emit_fixture_file "${FAKE_GRIDGRIND_TASK_PLAN_FILE:?}"
        ;;
    --print-protocol-catalog)
        if [[ "${2:-}" == '--full' ]]; then
            emit_fixture_file "${FAKE_GRIDGRIND_CATALOG_FILE:?}"
        else
            emit_fixture_file "${FAKE_GRIDGRIND_CATALOG_INDEX_FILE:?}"
        fi
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
    local help_overview_text=$2
    local help_protocol_text=$3
    local help_guidance_text=$4
    local catalog_text=$5
    local example_catalog_text=$6
    local task_catalog_text=$7
    local task_plan_text=$8
    local task_keyword_match_text=$9
    local doctor_report_text=${10}
    local noargs_failure_text=${11}
    local interactive_noargs_failure_text=${12}
    local case_dir
    local help_overview_file
    local help_protocol_file
    local help_guidance_file
    local catalog_index_file
    local catalog_file
    local example_catalog_file
    local task_catalog_file
    local task_plan_file
    local task_keyword_match_file
    local request_template_file
    local doctor_report_file
    local noargs_failure_file
    local interactive_noargs_failure_file

    case_dir="$(next_case_dir)"
    help_overview_file="$(write_case_fixture "${case_dir}" 'help-overview.txt' "${help_overview_text}")"
    help_protocol_file="$(write_case_fixture "${case_dir}" 'help-protocol.txt' "${help_protocol_text}")"
    help_guidance_file="$(write_case_fixture "${case_dir}" 'help-guidance.txt' "${help_guidance_text}")"
    catalog_index_file="$(write_case_fixture "${case_dir}" 'protocol-catalog-index.json' "${success_catalog_index}")"
    catalog_file="$(write_case_fixture "${case_dir}" 'protocol-catalog.json' "${catalog_text}")"
    example_catalog_file="$(write_case_fixture "${case_dir}" 'example-catalog.json' "${example_catalog_text}")"
    task_catalog_file="$(write_case_fixture "${case_dir}" 'task-catalog.json' "${task_catalog_text}")"
    task_plan_file="$(write_case_fixture "${case_dir}" 'task-plan.json' "${task_plan_text}")"
    task_keyword_match_file="$(write_case_fixture "${case_dir}" 'task-keyword-match.json' "${task_keyword_match_text}")"
    request_template_file="$(write_case_fixture "${case_dir}" 'request-template.json' "${success_request_template}")"
    doctor_report_file="$(write_case_fixture "${case_dir}" 'doctor-report.json' "${doctor_report_text}")"
    noargs_failure_file="$(write_case_fixture "${case_dir}" 'noargs-failure.json' "${noargs_failure_text}")"
    interactive_noargs_failure_file="$(
        write_case_fixture "${case_dir}" 'interactive-noargs-failure.json' "${interactive_noargs_failure_text}"
    )"

    if [[ "${unset_tmpdir}" == 'true' ]]; then
        env -u TMPDIR \
            FAKE_GRIDGRIND_HELP_OVERVIEW_FILE="${help_overview_file}" \
            FAKE_GRIDGRIND_HELP_PROTOCOL_FILE="${help_protocol_file}" \
            FAKE_GRIDGRIND_HELP_GUIDANCE_FILE="${help_guidance_file}" \
            FAKE_GRIDGRIND_CATALOG_INDEX_FILE="${catalog_index_file}" \
            FAKE_GRIDGRIND_CATALOG_FILE="${catalog_file}" \
            FAKE_GRIDGRIND_EXAMPLE_CATALOG_FILE="${example_catalog_file}" \
            FAKE_GRIDGRIND_TASK_CATALOG_FILE="${task_catalog_file}" \
            FAKE_GRIDGRIND_TASK_PLAN_FILE="${task_plan_file}" \
            FAKE_GRIDGRIND_TASK_KEYWORD_MATCH_FILE="${task_keyword_match_file}" \
            FAKE_GRIDGRIND_REQUEST_TEMPLATE_FILE="${request_template_file}" \
            FAKE_GRIDGRIND_DOCTOR_REPORT_FILE="${doctor_report_file}" \
            FAKE_GRIDGRIND_NOARGS_FAILURE_FILE="${noargs_failure_file}" \
            FAKE_GRIDGRIND_INTERACTIVE_NOARGS_FAILURE_FILE="${interactive_noargs_failure_file}" \
            "${verify_script}" binary "${fake_cli}" >/dev/null
        return 0
    fi

    FAKE_GRIDGRIND_HELP_OVERVIEW_FILE="${help_overview_file}" \
        FAKE_GRIDGRIND_HELP_PROTOCOL_FILE="${help_protocol_file}" \
        FAKE_GRIDGRIND_HELP_GUIDANCE_FILE="${help_guidance_file}" \
        FAKE_GRIDGRIND_CATALOG_INDEX_FILE="${catalog_index_file}" \
        FAKE_GRIDGRIND_CATALOG_FILE="${catalog_file}" \
        FAKE_GRIDGRIND_EXAMPLE_CATALOG_FILE="${example_catalog_file}" \
        FAKE_GRIDGRIND_TASK_CATALOG_FILE="${task_catalog_file}" \
        FAKE_GRIDGRIND_TASK_PLAN_FILE="${task_plan_file}" \
        FAKE_GRIDGRIND_TASK_KEYWORD_MATCH_FILE="${task_keyword_match_file}" \
        FAKE_GRIDGRIND_REQUEST_TEMPLATE_FILE="${request_template_file}" \
        FAKE_GRIDGRIND_DOCTOR_REPORT_FILE="${doctor_report_file}" \
        FAKE_GRIDGRIND_NOARGS_FAILURE_FILE="${noargs_failure_file}" \
        FAKE_GRIDGRIND_INTERACTIVE_NOARGS_FAILURE_FILE="${interactive_noargs_failure_file}" \
        "${verify_script}" binary "${fake_cli}" >/dev/null
}

run_verify_expect_success() {
    run_verify_with_fixture_texts \
        false \
        "${success_help_overview}" \
        "${success_help_protocol}" \
        "${success_help_guidance}" \
        "${success_catalog}" \
        "${success_example_catalog}" \
        "${success_task_catalog}" \
        "${success_task_plan}" \
        "${success_task_keyword_match}" \
        "${success_doctor_report}" \
        "${success_noargs_failure}" \
        "${success_noargs_failure}"
}

run_verify_expect_success_without_tmpdir() {
    run_verify_with_fixture_texts \
        true \
        "${success_help_overview}" \
        "${success_help_protocol}" \
        "${success_help_guidance}" \
        "${success_catalog}" \
        "${success_example_catalog}" \
        "${success_task_catalog}" \
        "${success_task_plan}" \
        "${success_task_keyword_match}" \
        "${success_doctor_report}" \
        "${success_noargs_failure}" \
        "${success_noargs_failure}"
}

run_verify_expect_failure() {
    local help_overview_text=$1
    local help_protocol_text=$2
    local help_guidance_text=$3
    local catalog_text=$4
    local example_catalog_text=${5:-${success_example_catalog}}
    local task_catalog_text=${6:-${success_task_catalog}}
    local task_plan_text=${7:-${success_task_plan}}
    local task_keyword_match_text=${8:-${success_task_keyword_match}}
    local doctor_report_text=${9:-${success_doctor_report}}
    local noargs_failure_text=${10:-${success_noargs_failure}}
    local interactive_noargs_failure_text=${11:-${noargs_failure_text}}
    if run_verify_with_fixture_texts \
        false \
        "${help_overview_text}" \
        "${help_protocol_text}" \
        "${help_guidance_text}" \
        "${catalog_text}" \
        "${example_catalog_text}" \
        "${task_catalog_text}" \
        "${task_plan_text}" \
        "${task_keyword_match_text}" \
        "${doctor_report_text}" \
        "${noargs_failure_text}" \
        "${interactive_noargs_failure_text}" >/dev/null 2>&1; then
        die "verifier unexpectedly succeeded"
    fi
}

run_verify_expect_success
run_verify_expect_success_without_tmpdir
run_verify_with_fixture_texts \
    false \
    "${success_help_overview}" \
    "${success_help_protocol}" \
    "${success_help_guidance}" \
    "${success_catalog}" \
    "${success_example_catalog}" \
    "${success_task_catalog}" \
    "${success_task_plan}" \
    "${success_task_keyword_match}" \
    "${success_doctor_report}" \
    "${success_noargs_failure}" \
    "$(append_fixture_line "${success_noargs_failure}" 'runtime-owned trailer')"

run_verify_expect_failure \
    "${success_help_overview}" \
    "$(append_fixture_line "${success_help_protocol}" 'FORCE_FORMULA_RECALC_ON_OPEN')" \
    "${success_help_guidance}" \
    "${success_catalog}"

run_verify_expect_failure \
    "${success_help_overview}" \
    "${success_help_protocol}" \
    "$(replace_fixture_token "${success_help_guidance}" 'WORKBOOK_HEALTH' 'WORKBOOK_HEALTH_BROKEN')" \
    "${success_catalog}"

run_verify_expect_failure \
    "${success_help_overview}" \
    "${success_help_protocol}" \
    "${success_help_guidance}" \
    "${success_catalog}" \
    "${success_example_catalog}" \
    "$(replace_fixture_token "${success_task_catalog}" 'SET_TABLE' 'NO_SUCH_MUTATION')"

run_verify_expect_failure \
    "${success_help_overview}" \
    "${success_help_protocol}" \
    "${success_help_guidance}" \
    "${success_catalog}" \
    "${success_example_catalog}" \
    "${success_task_catalog}" \
    "$(replace_fixture_token \
        "${success_task_plan}" \
        'generated-workbooks/dashboard.xlsx' \
        'generated-workbooks/dashboard.xls')"

run_verify_expect_failure \
    "${success_help_overview}" \
    "${success_help_protocol}" \
    "${success_help_guidance}" \
    "${success_catalog}" \
    "${success_example_catalog}" \
    "${success_task_catalog}" \
    "${success_task_plan}" \
    "${success_task_keyword_match}" \
    "$(replace_fixture_token \
        "${success_doctor_report}" \
        '"sourceType" : "NEW"' \
        '"sourceType" : "UTF8_FILE"')" \
    "${success_noargs_failure}"

run_verify_expect_failure \
    "${success_help_overview}" \
    "${success_help_protocol}" \
    "${success_help_guidance}" \
    "${success_catalog}" \
    "${success_example_catalog}" \
    "${success_task_catalog}" \
    "${success_task_plan}" \
    "$(replace_fixture_token \
        "${success_task_keyword_match}" \
        '"taskId" : "DASHBOARD"' \
        '"taskId" : "TABULAR_REPORT"')" \
    "${success_doctor_report}"

run_verify_expect_failure \
    "${success_help_overview}" \
    "${success_help_protocol}" \
    "${success_help_guidance}" \
    "${success_catalog}" \
    "${success_example_catalog}" \
    "${success_task_catalog}" \
    "${success_task_plan}" \
    "${success_task_keyword_match}" \
    "${success_doctor_report}" \
    "${success_noargs_failure}" \
    "$(replace_fixture_token \
        "${success_noargs_failure}" \
        '"command" : "execute"' \
        '"command" : "doctor-request"')"

printf 'verify-cli-contract regression: success\n'
