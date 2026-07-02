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

# shellcheck source=/dev/null
source "${repo_root}/scripts/lib/test-cli-case-file-support.sh"
# shellcheck source=/dev/null
source "${repo_root}/scripts/lib/test-cli-contract-regression-support.sh"

test_root="$(mktemp -d)"
cleanup() {
    rm -rf "${test_root}"
}
trap cleanup EXIT

readonly fake_cli="${test_root}/gridgrind"
write_fake_gridgrind_cli "${fake_cli}"

source "${repo_root}/scripts/lib/test-cli-contract-fixtures.sh"
load_test_cli_contract_fixtures

verify_case_counter=0

run_verify_expect_success() {
    run_fake_gridgrind_verify_with_fixture_texts \
        false \
        "${success_help_overview}" \
        "${success_help_protocol}" \
        "${success_help_guidance}" \
        "${success_source_types}" \
        "${success_persistence_types}" \
        "${success_step_types}" \
        "${success_mutation_action_types}" \
        "${success_assertion_types}" \
        "${success_inspection_query_types}" \
        "${success_execution_mode_types}" \
        "${success_execution_policy_input_type}" \
        "${success_example_catalog}" \
        "${success_task_catalog}" \
        "${success_task_plan}" \
        "${success_task_keyword_match}" \
        "${success_doctor_report}" \
        "${success_noargs_failure}" \
        "${success_noargs_failure}"
}

run_verify_expect_success_without_tmpdir() {
    run_fake_gridgrind_verify_with_fixture_texts \
        true \
        "${success_help_overview}" \
        "${success_help_protocol}" \
        "${success_help_guidance}" \
        "${success_source_types}" \
        "${success_persistence_types}" \
        "${success_step_types}" \
        "${success_mutation_action_types}" \
        "${success_assertion_types}" \
        "${success_inspection_query_types}" \
        "${success_execution_mode_types}" \
        "${success_execution_policy_input_type}" \
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
    local source_types_text=$4
    local persistence_types_text=${5:-${success_persistence_types}}
    local step_types_text=${6:-${success_step_types}}
    local mutation_action_types_text=${7:-${success_mutation_action_types}}
    local assertion_types_text=${8:-${success_assertion_types}}
    local inspection_query_types_text=${9:-${success_inspection_query_types}}
    local execution_mode_types_text=${10:-${success_execution_mode_types}}
    local execution_policy_input_type_text=${11:-${success_execution_policy_input_type}}
    local example_catalog_text=${12:-${success_example_catalog}}
    local task_catalog_text=${13:-${success_task_catalog}}
    local task_plan_text=${14:-${success_task_plan}}
    local task_keyword_match_text=${15:-${success_task_keyword_match}}
    local doctor_report_text=${16:-${success_doctor_report}}
    local noargs_failure_text=${17:-${success_noargs_failure}}
    local interactive_noargs_failure_text=${18:-${noargs_failure_text}}
    if run_fake_gridgrind_verify_with_fixture_texts \
        false \
        "${help_overview_text}" \
        "${help_protocol_text}" \
        "${help_guidance_text}" \
        "${source_types_text}" \
        "${persistence_types_text}" \
        "${step_types_text}" \
        "${mutation_action_types_text}" \
        "${assertion_types_text}" \
        "${inspection_query_types_text}" \
        "${execution_mode_types_text}" \
        "${execution_policy_input_type_text}" \
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
run_fake_gridgrind_verify_with_fixture_texts \
    false \
    "${success_help_overview}" \
    "${success_help_protocol}" \
    "${success_help_guidance}" \
    "${success_source_types}" \
    "${success_persistence_types}" \
    "${success_step_types}" \
    "${success_mutation_action_types}" \
    "${success_assertion_types}" \
    "${success_inspection_query_types}" \
    "${success_execution_mode_types}" \
    "${success_execution_policy_input_type}" \
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
    "${success_source_types}"

run_verify_expect_failure \
    "${success_help_overview}" \
    "${success_help_protocol}" \
    "$(replace_fixture_token "${success_help_guidance}" 'WORKBOOK_HEALTH' 'WORKBOOK_HEALTH_BROKEN')" \
    "${success_source_types}"

run_verify_expect_failure \
    "${success_help_overview}" \
    "${success_help_protocol}" \
    "${success_help_guidance}" \
    "${success_source_types}" \
    "${success_persistence_types}" \
    "${success_step_types}" \
    "${success_mutation_action_types}" \
    "${success_assertion_types}" \
    "${success_inspection_query_types}" \
    "${success_execution_mode_types}" \
    "${success_execution_policy_input_type}" \
    "${success_example_catalog}" \
    "$(replace_fixture_token "${success_task_catalog}" 'SET_TABLE' 'NO_SUCH_MUTATION')"

run_verify_expect_failure \
    "${success_help_overview}" \
    "${success_help_protocol}" \
    "${success_help_guidance}" \
    "${success_source_types}" \
    "${success_persistence_types}" \
    "${success_step_types}" \
    "${success_mutation_action_types}" \
    "${success_assertion_types}" \
    "${success_inspection_query_types}" \
    "${success_execution_mode_types}" \
    "${success_execution_policy_input_type}" \
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
    "${success_source_types}" \
    "${success_persistence_types}" \
    "${success_step_types}" \
    "${success_mutation_action_types}" \
    "${success_assertion_types}" \
    "${success_inspection_query_types}" \
    "${success_execution_mode_types}" \
    "${success_execution_policy_input_type}" \
    "${success_example_catalog}" \
    "${success_task_catalog}" \
    "${success_task_plan}" \
    "${success_task_keyword_match}" \
    "$(replace_fixture_token \
        "${success_doctor_report}" \
        '"sourceType":"NEW"' \
        '"sourceType":"UTF8_FILE"')" \
    "${success_noargs_failure}"

run_verify_expect_failure \
    "${success_help_overview}" \
    "${success_help_protocol}" \
    "${success_help_guidance}" \
    "${success_source_types}" \
    "${success_persistence_types}" \
    "${success_step_types}" \
    "${success_mutation_action_types}" \
    "${success_assertion_types}" \
    "${success_inspection_query_types}" \
    "${success_execution_mode_types}" \
    "${success_execution_policy_input_type}" \
    "${success_example_catalog}" \
    "${success_task_catalog}" \
    "${success_task_plan}" \
    "$(replace_fixture_token \
        "${success_task_keyword_match}" \
        '"taskId":"DASHBOARD"' \
        '"taskId":"TABULAR_REPORT"')" \
    "${success_doctor_report}"

run_verify_expect_failure \
    "${success_help_overview}" \
    "${success_help_protocol}" \
    "${success_help_guidance}" \
    "${success_source_types}" \
    "${success_persistence_types}" \
    "${success_step_types}" \
    "${success_mutation_action_types}" \
    "${success_assertion_types}" \
    "${success_inspection_query_types}" \
    "${success_execution_mode_types}" \
    "${success_execution_policy_input_type}" \
    "${success_example_catalog}" \
    "${success_task_catalog}" \
    "${success_task_plan}" \
    "${success_task_keyword_match}" \
    "${success_doctor_report}" \
    "${success_noargs_failure}" \
    "$(replace_fixture_token \
        "${success_noargs_failure}" \
        '"command":"execute"' \
        '"command":"doctor-request"')"

printf 'verify-cli-contract regression: success\n'
