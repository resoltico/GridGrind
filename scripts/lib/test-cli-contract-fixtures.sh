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
  "execution": {
    "mode": {"type": "FULL_XSSF"},
    "journal": { "level": "NORMAL" },
    "calculation": {
      "strategy": { "type": "DO_NOT_CALCULATE" },
      "markRecalculateOnOpen": false
    }
  },
  "formulaEnvironment": {
    "externalWorkbooks": [],
    "missingWorkbookPolicy": "ERROR",
    "udfToolpacks": []
  },
  "steps": []
}
JSON
}

load_test_cli_contract_fixtures() {
    local jar_path
    # shellcheck source=/dev/null
    source "${repo_root}/scripts/lib/cli-shadow-jar-support.sh"
    jar_path="$(ensure_cli_shadow_jar "${repo_root}")"

    success_help_overview="$(
        java -jar "${jar_path}" --help | tr -d '\r'
    )"
    success_help_protocol="$(
        java -jar "${jar_path}" --help-protocol | tr -d '\r'
    )"
    success_help_guidance="$(
        java -jar "${jar_path}" --help-guidance | tr -d '\r'
    )"
    success_catalog="$(
        java -jar "${jar_path}" --print-protocol-catalog | tr -d '\r'
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
        printf '%s' "${success_request_template}" | java -jar "${jar_path}" --doctor-request | tr -d '\r'
    )"
    success_noargs_failure="$(
        tmp_stdout="$(mktemp)"
        tmp_stderr="$(mktemp)"
        set +e
        java -jar "${jar_path}" >"${tmp_stdout}" 2>"${tmp_stderr}"
        exit_code=$?
        set -e
        [[ ${exit_code} -eq 2 ]] || {
            printf 'error: expected no-arg exit code 2, got %s\n' "${exit_code}" >&2
            rm -f "${tmp_stdout}" "${tmp_stderr}"
            return 1
        }
        [[ ! -s "${tmp_stderr}" ]] || {
            printf 'error: expected empty stderr for no-arg failure fixture\n' >&2
            rm -f "${tmp_stdout}" "${tmp_stderr}"
            return 1
        }
        tr -d '\r' < "${tmp_stdout}"
        rm -f "${tmp_stdout}" "${tmp_stderr}"
    )"
}
