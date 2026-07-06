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

# shellcheck source=/dev/null
source "${repo_root}/scripts/lib/test-cli-case-file-support.sh"
# shellcheck source=/dev/null
source "${repo_root}/scripts/lib/test-container-publication-support.sh"

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
readonly fake_discovery_verify="${test_root}/fake-verify-cli-discovery-execution.sh"
mkdir -p "${fake_bin}"
write_fake_docker_cli "${fake_bin}/docker"
write_fake_discovery_verify_script "${fake_discovery_verify}"

verify_case_counter=0

run_verify_expect_success() {
    run_fake_docker_verify_with_fixture_texts "$@"
}

run_verify_expect_failure() {
    local version_output=$1
    local latest_version_output=$2
    local help_overview_output=$3
    local help_protocol_output=$4
    local help_guidance_output=$5
    local catalog_index_output=$6
    local source_types_output=${7:-${success_source_types}}
    local persistence_types_output=${8:-${success_persistence_types}}
    local step_types_output=${9:-${success_step_types}}
    local mutation_action_types_output=${10:-${success_mutation_action_types}}
    local assertion_types_output=${11:-${success_assertion_types}}
    local inspection_query_types_output=${12:-${success_inspection_query_types}}
    local execution_mode_types_output=${13:-${success_execution_mode_types}}
    local execution_policy_input_type_output=${14:-${success_execution_policy_input_type}}
    local recipe_catalog_output=${15:-${success_recipe_catalog}}
    local recipe_request_output=${16:-${success_recipe_request}}
    local recipe_keyword_match_output=${17:-${success_recipe_keyword_match}}
    local doctor_report_output=${18:-${success_doctor_report}}
    local noargs_failure_output=${19:-${success_noargs_failure}}
    if run_fake_docker_verify_with_fixture_texts \
        "${version_output}" \
        "${latest_version_output}" \
        "${help_overview_output}" \
        "${help_protocol_output}" \
        "${help_guidance_output}" \
        "${catalog_index_output}" \
        "${source_types_output}" \
        "${persistence_types_output}" \
        "${step_types_output}" \
        "${mutation_action_types_output}" \
        "${assertion_types_output}" \
        "${inspection_query_types_output}" \
        "${execution_mode_types_output}" \
        "${execution_policy_input_type_output}" \
        "${recipe_catalog_output}" \
        "${recipe_request_output}" \
        "${recipe_keyword_match_output}" \
        "${doctor_report_output}" \
        "${noargs_failure_output}" >/dev/null 2>&1; then
        die "verifier unexpectedly succeeded"
    fi
}

expected_header="$(printf 'GridGrind 9.9.9\n%s' "${expected_description}")"
source "${repo_root}/scripts/lib/test-cli-contract-fixtures.sh"
load_test_cli_contract_fixtures

: > "${fake_log}"
run_verify_expect_success \
    "${expected_header}" \
    "${expected_header}" \
    "${success_help_overview}" \
    "${success_help_protocol}" \
    "${success_help_guidance}" \
    "${success_catalog_index}" \
    "${success_source_types}" \
    "${success_persistence_types}" \
    "${success_step_types}" \
    "${success_mutation_action_types}" \
    "${success_assertion_types}" \
    "${success_inspection_query_types}" \
    "${success_execution_mode_types}" \
    "${success_execution_policy_input_type}" \
    "${success_recipe_catalog}" \
    "${success_recipe_request}" \
    "${success_recipe_keyword_match}" \
    "${success_doctor_report}" \
    "${success_noargs_failure}"
grep -Fq 'pull ghcr.io/example/gridgrind:9.9.9' "${fake_log}" || die "verifier did not pull the version tag"
grep -Fq 'pull ghcr.io/example/gridgrind:latest' "${fake_log}" || die "verifier did not pull the latest tag"
grep -Fq 'run --rm ghcr.io/example/gridgrind:9.9.9 --help' "${fake_log}" || die \
    "verifier did not inspect the version tag help surface"
grep -Fq 'run --rm ghcr.io/example/gridgrind:latest --print-protocol-catalog' "${fake_log}" || die \
    "verifier did not inspect the latest tag catalog surface"
grep -Fq 'run --rm ghcr.io/example/gridgrind:latest --print-protocol-catalog --lookup mutationActionTypes' "${fake_log}" || die \
    "verifier did not inspect the latest tag scoped catalog group surface"
grep -Fq 'run --rm ghcr.io/example/gridgrind:latest --print-protocol-catalog --lookup nestedTypes:executionModeTypes' "${fake_log}" || die \
    "verifier did not inspect the latest tag nested catalog group surface"
grep -Fq 'run --rm ghcr.io/example/gridgrind:latest --print-protocol-catalog --lookup plainTypes:executionPolicyInputType' "${fake_log}" || die \
    "verifier did not inspect the latest tag plain catalog group surface"
grep -Fq 'run --rm ghcr.io/example/gridgrind:latest --print-recipe-catalog' "${fake_log}" || die \
    "verifier did not inspect the latest tag recipe-catalog surface"
grep -Fq 'run --rm ghcr.io/example/gridgrind:latest --print-recipe-catalog --lookup BUDGET' "${fake_log}" || die \
    "verifier did not inspect the latest tag example recipe-catalog detail surface"
grep -Fq 'run --rm ghcr.io/example/gridgrind:latest --print-recipe-catalog --lookup TABULAR_REPORT' "${fake_log}" || die \
    "verifier did not inspect the latest tag task recipe-catalog detail surface"
grep -Fq 'run --rm ghcr.io/example/gridgrind:latest --print-recipe --lookup DASHBOARD' "${fake_log}" || die \
    "verifier did not inspect the latest tag recipe surface"
grep -Fq 'run --rm ghcr.io/example/gridgrind:latest --print-recipe-keyword-match --query monthly sales dashboard with charts' "${fake_log}" || die \
    "verifier did not inspect the latest tag recipe-keyword-match surface"
grep -Fq 'discovery docker-image ghcr.io/example/gridgrind:9.9.9' "${fake_log}" || die \
    "verifier did not run discovery execution against the version tag"
grep -Fq 'discovery docker-image ghcr.io/example/gridgrind:latest' "${fake_log}" || die \
    "verifier did not run discovery execution against the latest tag"
grep -Fq 'run --rm -i -t ghcr.io/example/gridgrind:latest' "${fake_log}" || die \
    "verifier did not inspect the latest tag interactive no-arg failure surface"
grep -Fq 'run --rm -i ghcr.io/example/gridgrind:latest --doctor-request' "${fake_log}" || die \
    "verifier did not inspect the latest tag doctor surface"

FAKE_DISCOVERY_SHOULD_FAIL=1 run_verify_expect_failure \
    "${expected_header}" \
    "${expected_header}" \
    "${success_help_overview}" \
    "${success_help_protocol}" \
    "${success_help_guidance}" \
    "${success_catalog_index}" \
    "${success_source_types}" \
    "${success_persistence_types}" \
    "${success_step_types}" \
    "${success_mutation_action_types}" \
    "${success_assertion_types}" \
    "${success_inspection_query_types}" \
    "${success_execution_mode_types}" \
    "${success_execution_policy_input_type}" \
    "${success_recipe_catalog}" \
    "${success_recipe_request}" \
    "${success_recipe_keyword_match}" \
    "${success_doctor_report}" \
    "${success_noargs_failure}"
unset FAKE_DISCOVERY_SHOULD_FAIL

run_verify_expect_failure \
    "$(printf 'gridgrind 9.9.9')" \
    "$(printf 'gridgrind 9.9.9')" \
    "${success_help_overview}" \
    "${success_help_protocol}" \
    "${success_help_guidance}" \
    "${success_catalog_index}"
run_verify_expect_failure "$(printf 'GridGrind 9.9.9\nWrong description')" \
    "$(printf 'GridGrind 9.9.9\nWrong description')" \
    "${success_help_overview}" \
    "${success_help_protocol}" \
    "${success_help_guidance}" \
    "${success_catalog_index}"
run_verify_expect_failure \
    "${expected_header}" \
    "${expected_header}" \
    "${success_help_overview}" \
    "$(append_fixture_line "${success_help_protocol}" 'FORCE_FORMULA_RECALC_ON_OPEN')" \
    "${success_help_guidance}" \
    "${success_catalog_index}"
run_verify_expect_failure \
    "${expected_header}" \
    "${expected_header}" \
    "${success_help_overview}" \
    "${success_help_protocol}" \
    "${success_help_guidance}" \
    "$(replace_fixture_token \
        "${success_catalog_index}" \
        '"lookupNamespaces"' \
        '"lookupNamespacez"')"
run_verify_expect_failure \
    "${expected_header}" \
    "${expected_header}" \
    "${success_help_overview}" \
    "${success_help_protocol}" \
    "${success_help_guidance}" \
    "${success_catalog_index}" \
    "${success_source_types}" \
    "${success_persistence_types}" \
    "${success_step_types}" \
    "$(replace_fixture_token \
        "${success_mutation_action_types}" \
        '"SET_TABLE"' \
        '"SET_TABLEx"')"
run_verify_expect_failure \
    "${expected_header}" \
    "${expected_header}" \
    "${success_help_overview}" \
    "${success_help_protocol}" \
    "${success_help_guidance}" \
    "${success_catalog_index}" \
    "${success_source_types}" \
    "${success_persistence_types}" \
    "${success_step_types}" \
    "${success_mutation_action_types}" \
    "${success_assertion_types}" \
    "${success_inspection_query_types}" \
    "${success_execution_mode_types}" \
    "${success_execution_policy_input_type}" \
    "${success_recipe_catalog}" \
    "${success_recipe_request}" \
    "${success_recipe_keyword_match}" \
    "$(replace_fixture_token \
        "${success_doctor_report}" \
        '"valid":true' \
        '"valid":false')"
run_verify_expect_failure \
    "${expected_header}" \
    "${expected_header}" \
    "${success_help_overview}" \
    "${success_help_protocol}" \
    "${success_help_guidance}" \
    "${success_catalog_index}" \
    "${success_source_types}" \
    "${success_persistence_types}" \
    "${success_step_types}" \
    "${success_mutation_action_types}" \
    "${success_assertion_types}" \
    "${success_inspection_query_types}" \
    "${success_execution_mode_types}" \
    "$(replace_fixture_token \
        "${success_execution_policy_input_type}" \
        'execution.journal' \
        'executionJournal')"
run_verify_expect_failure \
    "${expected_header}" \
    "${expected_header}" \
    "${success_help_overview}" \
    "${success_help_protocol}" \
    "${success_help_guidance}" \
    "${success_catalog_index}" \
    "${success_source_types}" \
    "${success_persistence_types}" \
    "${success_step_types}" \
    "${success_mutation_action_types}" \
    "${success_assertion_types}" \
    "${success_inspection_query_types}" \
    "${success_execution_mode_types}" \
    "${success_execution_policy_input_type}" \
    "${success_recipe_catalog}" \
    "${success_recipe_request}" \
    "$(replace_fixture_token \
        "${success_recipe_keyword_match}" \
        '"recipeId":"DASHBOARD"' \
        '"recipeId":"TABULAR_REPORT"')"

printf 'verify-container-publication regression: success\n'
