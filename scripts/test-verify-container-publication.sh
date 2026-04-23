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
                printf '%s' "${FAKE_DOCKER_HELP_OUTPUT:?}"
                ;;
            --version)
                case "${image_ref}" in
                    *:latest)
                        printf '%s' "${FAKE_DOCKER_LATEST_VERSION_OUTPUT:?}"
                        ;;
                    *)
                        printf '%s' "${FAKE_DOCKER_VERSION_OUTPUT:?}"
                        ;;
                esac
                ;;
            --help)
                printf '%s' "${FAKE_DOCKER_HELP_OUTPUT:?}"
                ;;
            --print-task-catalog)
                printf '%s' "${FAKE_DOCKER_TASK_CATALOG_OUTPUT:?}"
                ;;
            --print-task-plan)
                printf '%s' "${FAKE_DOCKER_TASK_PLAN_OUTPUT:?}"
                ;;
            --print-goal-plan)
                printf '%s' "${FAKE_DOCKER_GOAL_PLAN_OUTPUT:?}"
                ;;
            --doctor-request)
                printf '%s' "${FAKE_DOCKER_DOCTOR_REPORT_OUTPUT:?}"
                ;;
            --print-protocol-catalog)
                printf '%s' "${FAKE_DOCKER_CATALOG_OUTPUT:?}"
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

run_verify_expect_success() {
    PATH="${fake_bin}:${PATH}" \
        FAKE_DOCKER_LOG="${fake_log}" \
        FAKE_DOCKER_VERSION_OUTPUT="$1" \
        FAKE_DOCKER_LATEST_VERSION_OUTPUT="$2" \
        FAKE_DOCKER_HELP_OUTPUT="$3" \
        FAKE_DOCKER_CATALOG_OUTPUT="$4" \
        FAKE_DOCKER_TASK_CATALOG_OUTPUT="$5" \
        FAKE_DOCKER_TASK_PLAN_OUTPUT="$6" \
        FAKE_DOCKER_GOAL_PLAN_OUTPUT="$7" \
        FAKE_DOCKER_DOCTOR_REPORT_OUTPUT="$8" \
        GRIDGRIND_PUBLICATION_VERIFY_RETRIES=1 \
        GRIDGRIND_PUBLICATION_VERIFY_DELAY_SECONDS=0 \
        "${verify_script}" "ghcr.io/example/gridgrind" "9.9.9" >/dev/null
}

run_verify_expect_failure() {
    if PATH="${fake_bin}:${PATH}" \
        FAKE_DOCKER_LOG="${fake_log}" \
        FAKE_DOCKER_VERSION_OUTPUT="$1" \
        FAKE_DOCKER_LATEST_VERSION_OUTPUT="$2" \
        FAKE_DOCKER_HELP_OUTPUT="$3" \
        FAKE_DOCKER_CATALOG_OUTPUT="$4" \
        FAKE_DOCKER_TASK_CATALOG_OUTPUT="$5" \
        FAKE_DOCKER_TASK_PLAN_OUTPUT="$6" \
        FAKE_DOCKER_GOAL_PLAN_OUTPUT="$7" \
        FAKE_DOCKER_DOCTOR_REPORT_OUTPUT="$8" \
        GRIDGRIND_PUBLICATION_VERIFY_RETRIES=1 \
        GRIDGRIND_PUBLICATION_VERIFY_DELAY_SECONDS=0 \
        "${verify_script}" "ghcr.io/example/gridgrind" "9.9.9" >/dev/null 2>&1; then
        die "verifier unexpectedly succeeded"
    fi
}

expected_header="$(printf 'GridGrind 9.9.9\n%s' "${expected_description}")"
source "${repo_root}/scripts/lib/test-cli-contract-fixtures.sh"

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
    "${success_help/markRecalculateOnOpen=true/FORCE_FORMULA_RECALC_ON_OPEN}" \
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
    "${success_doctor_report/\"valid\": true/\"valid\": false}"
run_verify_expect_failure \
    "${expected_header}" \
    "${expected_header}" \
    "${success_help}" \
    "${success_catalog}" \
    "${success_task_catalog}" \
    "${success_task_plan}" \
    "${success_goal_plan/\"id\": \"DASHBOARD\"/\"id\": \"TABULAR_REPORT\"}" \
    "${success_doctor_report/\"valid\": true/\"valid\": false}"

printf 'verify-container-publication regression: success\n'
