#!/usr/bin/env bash
# Guard the lazy, configuration-cache-safe Jazzer tool-task surface. The verification tasks must
# stay listable without unrelated Gradle properties, and the common read-only tool commands must
# keep storing or reusing the configuration cache instead of discarding it.

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

run_and_capture() {
    local output_path=$1
    shift
    local command_pid=0
    local heartbeat_pid=0
    "$@" >"${output_path}" 2>&1 &
    command_pid=$!
    (
        while kill -0 "${command_pid}" 2>/dev/null; do
            sleep 15
            if kill -0 "${command_pid}" 2>/dev/null; then
                printf 'jazzer-tool-task-cache regression: waiting for %s...\n' "$*"
            fi
        done
    ) &
    heartbeat_pid=$!
    if ! wait "${command_pid}"; then
        wait "${heartbeat_pid}" 2>/dev/null || true
        cat "${output_path}" >&2
        die "command failed: $*"
    fi
    wait "${heartbeat_pid}" 2>/dev/null || true
}

require_log_contains() {
    local output_path=$1
    local needle=$2
    grep -Fq "${needle}" "${output_path}" || {
        cat "${output_path}" >&2
        die "expected ${needle} in ${output_path}"
    }
}

require_log_contains_any() {
    local output_path=$1
    shift
    local needle=''
    for needle in "$@"; do
        if grep -Fq "${needle}" "${output_path}"; then
            return 0
        fi
    done
    cat "${output_path}" >&2
    die "expected one of [$*] in ${output_path}"
}

require_log_omits() {
    local output_path=$1
    local needle=$2
    if grep -Fq "${needle}" "${output_path}"; then
        cat "${output_path}" >&2
        die "unexpected ${needle} in ${output_path}"
    fi
}

readonly script_dir="$(resolve_script_dir)"
readonly repo_root="$(cd -P -- "${script_dir}/.." && pwd)"
readonly jazzer_project_dir="${repo_root}/jazzer"
readonly replay_input_path="${repo_root}/jazzer/src/fuzz/resources/dev/erst/gridgrind/jazzer/protocol/ProtocolRequestFuzzTestInputs/readRequest/data_validation_request.json"
readonly replay_expectation_path="${repo_root}/jazzer/src/fuzz/resources/dev/erst/gridgrind/jazzer/promoted-metadata/protocol-request/data_validation_request.txt"
readonly tmp_dir="$(mktemp -d)"
readonly tasks_log="${tmp_dir}/tasks.log"
readonly status_first_log="${tmp_dir}/status-first.log"
readonly status_second_log="${tmp_dir}/status-second.log"
readonly replay_log="${tmp_dir}/replay.log"
readonly replay_expected_outcome="$(grep -F 'Outcome:' "${replay_expectation_path}" | head -n 1)"
readonly replay_expected_decode_outcome="$(grep -F 'Decode Outcome:' "${replay_expectation_path}" | head -n 1)"

cleanup() {
    rm -rf "${tmp_dir}"
}
trap cleanup EXIT

run_and_capture \
    "${tasks_log}" \
    "${repo_root}/gradlew" \
    --project-dir "${jazzer_project_dir}" \
    tasks \
    --group verification \
    --console=plain \
    --no-daemon
require_log_contains "${tasks_log}" "jazzerSummarizeRun - Builds the latest summary artifacts for one completed Jazzer run."
require_log_contains "${tasks_log}" "jazzerReplay - Replays one local input against a single Jazzer harness."
require_log_omits "${tasks_log}" "Missing Gradle property"
require_log_omits "${tasks_log}" "Configuration cache entry discarded"

run_and_capture \
    "${status_first_log}" \
    "${repo_root}/gradlew" \
    --project-dir "${jazzer_project_dir}" \
    -PjazzerTarget=protocol-request \
    jazzerStatus \
    --console=plain \
    --no-daemon
require_log_contains "${status_first_log}" "Jazzer Status"
require_log_contains_any \
    "${status_first_log}" \
    "Configuration cache entry stored." \
    "Configuration cache entry reused."
require_log_omits "${status_first_log}" "Configuration cache entry discarded"

run_and_capture \
    "${status_second_log}" \
    "${repo_root}/gradlew" \
    --project-dir "${jazzer_project_dir}" \
    -PjazzerTarget=protocol-request \
    jazzerStatus \
    --console=plain \
    --no-daemon
require_log_contains "${status_second_log}" "Jazzer Status"
require_log_contains "${status_second_log}" "Configuration cache entry reused."
require_log_omits "${status_second_log}" "Configuration cache entry discarded"

run_and_capture \
    "${replay_log}" \
    "${repo_root}/gradlew" \
    --project-dir "${jazzer_project_dir}" \
    -PjazzerTarget=protocol-request \
    -PjazzerInput="${replay_input_path}" \
    jazzerReplay \
    --console=plain \
    --no-daemon
require_log_contains "${replay_log}" "${replay_expected_outcome}"
require_log_contains "${replay_log}" "${replay_expected_decode_outcome}"
require_log_contains_any \
    "${replay_log}" \
    "Configuration cache entry stored." \
    "Configuration cache entry reused."
require_log_omits "${replay_log}" "jsonOutput"
require_log_omits "${replay_log}" "Configuration cache entry discarded"

printf 'jazzer-tool-task-cache regression: success\n'
