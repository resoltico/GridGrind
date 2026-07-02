#!/usr/bin/env bash
# Canonical public-contract payloads derived from the built GridGrind CLI artifact.

replace_fixture_token() {
    local text=$1
    local needle=$2
    local replacement=$3
    [[ "${needle}" != "${replacement}" ]] || {
        printf 'error: replacement must differ from fixture token %s\n' "${needle}" >&2
        return 1
    }
    [[ "${text}" == *"${needle}"* ]] || {
        printf 'error: expected fixture token not found: %s\n' "${needle}" >&2
        return 1
    }
    printf '%s' "${text//${needle}/${replacement}}"
}

append_fixture_line() {
    local text=$1
    local appended_line=$2
    printf '%s\n%s' "${text}" "${appended_line}"
}

print_cli_contract_minimal_request() {
    cat <<'JSON'
{
  "protocolVersion": "V1",
  "source": { "type": "NEW" },
  "persistence": { "type": "NONE" },
  "steps": []
}
JSON
}

load_test_cli_contract_fixtures() {
    local jar_path
    local doctor_execution_root
    # shellcheck source=/dev/null
    source "${repo_root}/scripts/lib/cli-shadow-jar-support.sh"
    jar_path="$(ensure_cli_shadow_jar "${repo_root}")"
    mkdir -p "${repo_root}/tmp"
    doctor_execution_root="$(mktemp -d "${repo_root}/tmp/test-cli-contract-fixtures.XXXXXX")"

    success_help_overview="$(
        java -jar "${jar_path}" --help | tr -d '\r'
    )"
    success_help_protocol="$(
        java -jar "${jar_path}" --help-protocol | tr -d '\r'
    )"
    success_help_guidance="$(
        java -jar "${jar_path}" --help-guidance | tr -d '\r'
    )"
    success_catalog_index="$(
        java -jar "${jar_path}" --print-protocol-catalog | tr -d '\r'
    )"
    success_source_types="$(
        java -jar "${jar_path}" --print-protocol-catalog --lookup sourceTypes | tr -d '\r'
    )"
    success_persistence_types="$(
        java -jar "${jar_path}" --print-protocol-catalog --lookup persistenceTypes | tr -d '\r'
    )"
    success_step_types="$(
        java -jar "${jar_path}" --print-protocol-catalog --lookup stepTypes | tr -d '\r'
    )"
    success_mutation_action_types="$(
        java -jar "${jar_path}" --print-protocol-catalog --lookup mutationActionTypes | tr -d '\r'
    )"
    success_assertion_types="$(
        java -jar "${jar_path}" --print-protocol-catalog --lookup assertionTypes | tr -d '\r'
    )"
    success_inspection_query_types="$(
        java -jar "${jar_path}" --print-protocol-catalog --lookup inspectionQueryTypes | tr -d '\r'
    )"
    success_execution_mode_types="$(
        java -jar "${jar_path}" --print-protocol-catalog --lookup nestedTypes:executionModeTypes | tr -d '\r'
    )"
    success_execution_policy_input_type="$(
        java -jar "${jar_path}" --print-protocol-catalog --lookup plainTypes:executionPolicyInputType | tr -d '\r'
    )"
    success_example_catalog="$(
        java -jar "${jar_path}" --print-example-catalog | tr -d '\r'
    )"
    success_task_catalog="$(
        java -jar "${jar_path}" --print-task-catalog | tr -d '\r'
    )"
    success_task_plan="$(
        java -jar "${jar_path}" --print-task-plan --lookup DASHBOARD | tr -d '\r'
    )"
    success_task_keyword_match="$(
        java -jar "${jar_path}" --print-task-keyword-match --query "monthly sales dashboard with charts" | tr -d '\r'
    )"
    success_request_template="$(
        java -jar "${jar_path}" --print-request-template | tr -d '\r'
    )"
    success_doctor_report="$(
        printf '%s' "${success_request_template}" \
            | java -jar "${jar_path}" --doctor-request --execution-root "${doctor_execution_root}" | tr -d '\r'
    )"
    success_noargs_failure="$(
        tmp_stdout="$(mktemp)"
        tmp_stderr="$(mktemp)"
        set +e
        java -jar "${jar_path}" < /dev/null >"${tmp_stdout}" 2>"${tmp_stderr}"
        exit_code=$?
        set -e
        [[ ${exit_code} -eq 2 ]] || {
            printf 'error: expected no-arg exit code 2, got %s\n' "${exit_code}" >&2
            rm -f "${tmp_stdout}" "${tmp_stderr}"
            rm -rf "${doctor_execution_root}"
            return 1
        }
        [[ ! -s "${tmp_stdout}" ]] || {
            printf 'error: expected empty stdout for no-arg failure fixture\n' >&2
            rm -f "${tmp_stdout}" "${tmp_stderr}"
            rm -rf "${doctor_execution_root}"
            return 1
        }
        tr -d '\r' < "${tmp_stderr}"
        rm -f "${tmp_stdout}" "${tmp_stderr}"
    )"
    rm -rf "${doctor_execution_root}"
}
