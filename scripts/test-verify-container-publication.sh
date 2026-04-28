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
mkdir -p "${fake_bin}"

cat > "${fake_bin}/docker" <<'EOF'
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

log_path=${FAKE_DOCKER_LOG:?}
printf '%s\n' "$*" >> "${log_path}"

args=("$@")
offset=0
if [[ ${#args[@]} -ge 2 && "${args[0]}" == "--config" ]]; then
    offset=2
fi

if [[ ${#args[@]} -le ${offset} ]]; then
    printf 'unexpected docker invocation: %s\n' "$*" >&2
    exit 1
fi

command=${args[${offset}]}
case "${command}" in
    pull)
        image_ref=${args[$((offset + 1))]:-}
        [[ -n "${image_ref}" ]] || exit 1
        exit 0
        ;;
    run)
        run_index=$((offset + 1))
        while [[ ${#args[@]} -gt ${run_index} ]]; do
            case "${args[${run_index}]}" in
                --rm|--interactive|--tty|-i|-t|-it|-ti)
                    ((run_index += 1))
                    ;;
                *)
                    break
                    ;;
            esac
        done
        image_ref=${args[${run_index}]:-}
        cli_flag=${args[$((run_index + 1))]:-}
        case "${cli_flag}" in
            '')
                emit_fixture_file "${FAKE_DOCKER_HELP_OUTPUT_FILE:?}"
                ;;
            --version)
                case "${image_ref}" in
                    *:latest)
                        emit_fixture_file "${FAKE_DOCKER_LATEST_VERSION_OUTPUT_FILE:?}"
                        ;;
                    *)
                        emit_fixture_file "${FAKE_DOCKER_VERSION_OUTPUT_FILE:?}"
                        ;;
                esac
                ;;
            --help)
                emit_fixture_file "${FAKE_DOCKER_HELP_OUTPUT_FILE:?}"
                ;;
            --print-task-catalog)
                emit_fixture_file "${FAKE_DOCKER_TASK_CATALOG_OUTPUT_FILE:?}"
                ;;
            --print-task-plan)
                emit_fixture_file "${FAKE_DOCKER_TASK_PLAN_OUTPUT_FILE:?}"
                ;;
            --print-goal-plan)
                emit_fixture_file "${FAKE_DOCKER_GOAL_PLAN_OUTPUT_FILE:?}"
                ;;
            --doctor-request)
                emit_fixture_file "${FAKE_DOCKER_DOCTOR_REPORT_OUTPUT_FILE:?}"
                ;;
            --print-protocol-catalog)
                emit_fixture_file "${FAKE_DOCKER_CATALOG_OUTPUT_FILE:?}"
                ;;
            *)
                printf 'unexpected docker run invocation: %s\n' "$*" >&2
                exit 1
                ;;
        esac
        ;;
    *)
        printf 'unexpected docker subcommand: %s\n' "${command}" >&2
        exit 1
        ;;
esac
EOF
chmod +x "${fake_bin}/docker"

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
    local version_output=$1
    local latest_version_output=$2
    local help_output=$3
    local catalog_output=$4
    local task_catalog_output=$5
    local task_plan_output=$6
    local goal_plan_output=$7
    local doctor_report_output=$8
    local case_dir
    local version_output_file
    local latest_version_output_file
    local help_output_file
    local catalog_output_file
    local task_catalog_output_file
    local task_plan_output_file
    local goal_plan_output_file
    local doctor_report_output_file

    case_dir="$(next_case_dir)"
    version_output_file="$(write_case_fixture "${case_dir}" 'version.txt' "${version_output}")"
    latest_version_output_file="$(write_case_fixture "${case_dir}" 'latest-version.txt' "${latest_version_output}")"
    help_output_file="$(write_case_fixture "${case_dir}" 'help.txt' "${help_output}")"
    catalog_output_file="$(write_case_fixture "${case_dir}" 'protocol-catalog.json' "${catalog_output}")"
    task_catalog_output_file="$(write_case_fixture "${case_dir}" 'task-catalog.json' "${task_catalog_output}")"
    task_plan_output_file="$(write_case_fixture "${case_dir}" 'task-plan.json' "${task_plan_output}")"
    goal_plan_output_file="$(write_case_fixture "${case_dir}" 'goal-plan.json' "${goal_plan_output}")"
    doctor_report_output_file="$(write_case_fixture "${case_dir}" 'doctor-report.json' "${doctor_report_output}")"

    PATH="${fake_bin}:${PATH}" \
        FAKE_DOCKER_LOG="${fake_log}" \
        FAKE_DOCKER_VERSION_OUTPUT_FILE="${version_output_file}" \
        FAKE_DOCKER_LATEST_VERSION_OUTPUT_FILE="${latest_version_output_file}" \
        FAKE_DOCKER_HELP_OUTPUT_FILE="${help_output_file}" \
        FAKE_DOCKER_CATALOG_OUTPUT_FILE="${catalog_output_file}" \
        FAKE_DOCKER_TASK_CATALOG_OUTPUT_FILE="${task_catalog_output_file}" \
        FAKE_DOCKER_TASK_PLAN_OUTPUT_FILE="${task_plan_output_file}" \
        FAKE_DOCKER_GOAL_PLAN_OUTPUT_FILE="${goal_plan_output_file}" \
        FAKE_DOCKER_DOCTOR_REPORT_OUTPUT_FILE="${doctor_report_output_file}" \
        GRIDGRIND_PUBLICATION_VERIFY_RETRIES=1 \
        GRIDGRIND_PUBLICATION_VERIFY_DELAY_SECONDS=0 \
        "${verify_script}" "ghcr.io/example/gridgrind" "9.9.9" >/dev/null
}

run_verify_expect_success() {
    run_verify_with_fixture_texts "$@"
}

run_verify_expect_failure() {
    if run_verify_with_fixture_texts "$@" >/dev/null 2>&1; then
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
    "${success_help}" \
    "${success_catalog}" \
    "${success_task_catalog}" \
    "${success_task_plan}" \
    "${success_goal_plan}" \
    "${success_doctor_report}"
grep -Fq 'pull ghcr.io/example/gridgrind:9.9.9' "${fake_log}" || die "verifier did not pull the version tag"
grep -Fq 'pull ghcr.io/example/gridgrind:latest' "${fake_log}" || die "verifier did not pull the latest tag"
grep -Fq 'run --rm ghcr.io/example/gridgrind:9.9.9 --help' "${fake_log}" || die \
    "verifier did not inspect the version tag help surface"
grep -Fq 'run --rm ghcr.io/example/gridgrind:latest --print-protocol-catalog' "${fake_log}" || die \
    "verifier did not inspect the latest tag catalog surface"
grep -Fq 'run --rm ghcr.io/example/gridgrind:latest --print-task-catalog' "${fake_log}" || die \
    "verifier did not inspect the latest tag task-catalog surface"
grep -Fq 'run --rm ghcr.io/example/gridgrind:latest --print-task-plan DASHBOARD' "${fake_log}" || die \
    "verifier did not inspect the latest tag task-plan surface"
grep -Fq 'run --rm ghcr.io/example/gridgrind:latest --print-goal-plan monthly sales dashboard with charts' "${fake_log}" || die \
    "verifier did not inspect the latest tag goal-plan surface"
grep -Fq 'run --rm -i -t ghcr.io/example/gridgrind:latest' "${fake_log}" || die \
    "verifier did not inspect the latest tag interactive no-arg help surface"
grep -Fq 'run --rm -i ghcr.io/example/gridgrind:latest --doctor-request' "${fake_log}" || die \
    "verifier did not inspect the latest tag doctor surface"

run_verify_expect_failure \
    "$(printf 'gridgrind 9.9.9')" \
    "$(printf 'gridgrind 9.9.9')" \
    "${success_help}" \
    "${success_catalog}" \
    "${success_task_catalog}" \
    "${success_task_plan}" \
    "${success_goal_plan}" \
    "${success_doctor_report}"
run_verify_expect_failure "$(printf 'GridGrind 9.9.9\nWrong description')" \
    "$(printf 'GridGrind 9.9.9\nWrong description')" \
    "${success_help}" \
    "${success_catalog}" \
    "${success_task_catalog}" \
    "${success_task_plan}" \
    "${success_goal_plan}" \
    "${success_doctor_report}"
run_verify_expect_failure \
    "${expected_header}" \
    "${expected_header}" \
    "$(append_fixture_line "${success_help}" 'FORCE_FORMULA_RECALC_ON_OPEN')" \
    "${success_catalog}" \
    "${success_task_catalog}" \
    "${success_task_plan}" \
    "${success_goal_plan}" \
    "${success_doctor_report}"
run_verify_expect_failure \
    "${expected_header}" \
    "${expected_header}" \
    "${success_help}" \
    "${success_catalog}" \
    "${success_task_catalog}" \
    "${success_task_plan}" \
    "${success_goal_plan}" \
    "$(replace_fixture_token \
        "${success_doctor_report}" \
        '"valid" : true' \
        '"valid" : false')"
run_verify_expect_failure \
    "${expected_header}" \
    "${expected_header}" \
    "${success_help}" \
    "${success_catalog}" \
    "${success_task_catalog}" \
    "${success_task_plan}" \
    "$(replace_fixture_token \
        "${success_goal_plan}" \
        '"id" : "DASHBOARD"' \
        '"id" : "TABULAR_REPORT"')" \
    "$(replace_fixture_token \
        "${success_doctor_report}" \
        '"valid" : true' \
        '"valid" : false')"

printf 'verify-container-publication regression: success\n'
