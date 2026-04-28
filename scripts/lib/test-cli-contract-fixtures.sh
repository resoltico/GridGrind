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

ensure_cli_shadow_jar() {
    local jar_path="${repo_root}/cli/build/libs/gridgrind.jar"
    if [[ -f "${jar_path}" ]]; then
        printf '%s\n' "${jar_path}"
        return 0
    fi
    "${repo_root}/gradlew" :cli:shadowJar --console=plain
    [[ -f "${jar_path}" ]] || {
        printf 'error: expected CLI shadow jar at %s after build\n' "${jar_path}" >&2
        return 1
    }
    printf '%s\n' "${jar_path}"
}

load_test_cli_contract_fixtures() {
    local jar_path
    jar_path="$(ensure_cli_shadow_jar)"

    success_help="$(
        java -jar "${jar_path}" --help | tr -d '\r'
    )"
    success_catalog="$(
        java -jar "${jar_path}" --print-protocol-catalog | tr -d '\r'
    )"
    success_task_catalog="$(
        java -jar "${jar_path}" --print-task-catalog | tr -d '\r'
    )"
    success_task_plan="$(
        java -jar "${jar_path}" --print-task-plan DASHBOARD | tr -d '\r'
    )"
    success_goal_plan="$(
        java -jar "${jar_path}" --print-goal-plan "monthly sales dashboard with charts" | tr -d '\r'
    )"
    success_doctor_report="$(
        printf '%s\n' '{"source":{"type":"NEW"},"steps":[]}' \
            | java -jar "${jar_path}" --doctor-request | tr -d '\r'
    )"
}
