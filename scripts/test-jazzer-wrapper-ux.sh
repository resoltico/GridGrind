#!/usr/bin/env bash
# Keep the Jazzer wrapper help and invalid-target surface project-owned instead of leaking Gradle.

set -euo pipefail

die() {
    printf 'error: %s\n' "$1" >&2
    exit 1
}

require_contains() {
    local text=$1
    local needle=$2
    local description=$3
    if ! grep -Fq -- "${needle}" <<<"${text}"; then
        die "${description}"
    fi
}

require_absent() {
    local text=$1
    local needle=$2
    local description=$3
    if grep -Fq -- "${needle}" <<<"${text}"; then
        die "${description}"
    fi
}

require_file_contains() {
    local file_path=$1
    local needle=$2
    local description=$3
    if ! grep -Fq -- "${needle}" "${file_path}"; then
        die "${description}"
    fi
}

require_file_absent() {
    local file_path=$1
    local needle=$2
    local description=$3
    if grep -Fq -- "${needle}" "${file_path}"; then
        die "${description}"
    fi
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
readonly temp_parent="${repo_root}/tmp/test-jazzer-wrapper-ux"

mkdir -p "${temp_parent}"
tmp_dir="$(mktemp -d "${temp_parent%/}/run.XXXXXX")"
cleanup() {
    rm -rf "${tmp_dir}"
}
trap cleanup EXIT

capture_output() {
    local output_path=$1
    shift
    "$@" 2>&1 | tee "${output_path}" >/dev/null
}
readonly status_help_path="${tmp_dir}/status-help.txt"
readonly protocol_help_path="${tmp_dir}/fuzz-protocol-request-help.txt"
readonly fuzz_all_help_path="${tmp_dir}/fuzz-all-help.txt"
readonly replay_help_path="${tmp_dir}/replay-help.txt"
readonly promote_help_path="${tmp_dir}/promote-help.txt"
readonly invalid_report_output_path="${tmp_dir}/invalid-report.txt"
readonly invalid_replay_target_output_path="${tmp_dir}/invalid-replay-target.txt"
readonly missing_replay_input_output_path="${tmp_dir}/missing-replay-input.txt"
readonly invalid_promote_target_output_path="${tmp_dir}/invalid-promote-target.txt"
readonly missing_promote_input_output_path="${tmp_dir}/missing-promote-input.txt"

capture_output "${status_help_path}" "${repo_root}/jazzer/bin/status" --help
require_file_contains "${status_help_path}" "Usage: jazzer/bin/status" \
    "status --help no longer prints project-owned wrapper usage"
require_file_absent "${status_help_path}" "USAGE: gradlew" \
    "status --help leaked Gradle help instead of wrapper usage"
require_file_absent "${status_help_path}" "Welcome to Gradle" \
    "status --help booted Gradle instead of returning wrapper usage"

capture_output "${protocol_help_path}" "${repo_root}/jazzer/bin/fuzz-protocol-request" --help
require_file_contains "${protocol_help_path}" "Usage: jazzer/bin/fuzz-protocol-request" \
    "fuzz-protocol-request --help no longer prints project-owned wrapper usage"
require_file_absent "${protocol_help_path}" "USAGE: gradlew" \
    "fuzz-protocol-request --help leaked Gradle help instead of wrapper usage"
require_file_absent "${protocol_help_path}" "Welcome to Gradle" \
    "fuzz-protocol-request --help booted Gradle instead of returning wrapper usage"

capture_output "${fuzz_all_help_path}" "${repo_root}/jazzer/bin/fuzz-all" --help
require_file_contains "${fuzz_all_help_path}" "Usage: jazzer/bin/fuzz-all" \
    "fuzz-all --help no longer prints project-owned wrapper usage"
require_file_absent "${fuzz_all_help_path}" "USAGE: gradlew" \
    "fuzz-all --help leaked Gradle help instead of wrapper usage"

capture_output "${replay_help_path}" "${repo_root}/jazzer/bin/replay" --help
require_file_contains "${replay_help_path}" "Usage: jazzer/bin/replay" \
    "replay --help no longer prints project-owned wrapper usage"
require_file_contains "${replay_help_path}" "Valid targets:" \
    "replay --help no longer lists valid replay targets"

capture_output "${promote_help_path}" "${repo_root}/jazzer/bin/promote" --help
require_file_contains "${promote_help_path}" "Usage: jazzer/bin/promote" \
    "promote --help no longer prints project-owned wrapper usage"
require_file_contains "${promote_help_path}" "Valid targets:" \
    "promote --help no longer lists valid promotion targets"

set +e
capture_output "${invalid_report_output_path}" \
    "${repo_root}/jazzer/bin/report" no-such-target --console=plain
invalid_report_exit=$?
set -e

[[ ${invalid_report_exit} -eq 2 ]] || die \
    "report no-such-target exited ${invalid_report_exit}; expected the wrapper to reject it with code 2"
require_file_contains "${invalid_report_output_path}" "Unknown Jazzer target: no-such-target" \
    "report no-such-target no longer explains the invalid target"
require_file_contains "${invalid_report_output_path}" "Usage: jazzer/bin/report" \
    "report no-such-target no longer prints wrapper usage guidance"
require_file_absent "${invalid_report_output_path}" "Exception in thread" \
    "report no-such-target leaked a Java stack trace"
require_file_absent "${invalid_report_output_path}" "Execution failed for task" \
    "report no-such-target leaked a Gradle task failure instead of a wrapper-level error"
require_file_absent "${invalid_report_output_path}" "Welcome to Gradle" \
    "report no-such-target booted Gradle instead of rejecting the target up front"

set +e
capture_output "${invalid_replay_target_output_path}" \
    "${repo_root}/jazzer/bin/replay" no-such-target /tmp/does-not-matter.bin
invalid_replay_target_exit=$?
set -e

[[ ${invalid_replay_target_exit} -eq 2 ]] || die \
    "replay no-such-target exited ${invalid_replay_target_exit}; expected wrapper rejection with code 2"
require_file_contains "${invalid_replay_target_output_path}" "Unknown Jazzer target: no-such-target" \
    "replay no-such-target no longer explains the invalid target"
require_file_contains "${invalid_replay_target_output_path}" "Usage: jazzer/bin/replay" \
    "replay no-such-target no longer prints replay usage guidance"
require_file_absent "${invalid_replay_target_output_path}" "Usage: ${repo_root}/jazzer/bin/_run-task" \
    "replay no-such-target fell back to the internal _run-task usage surface"
require_file_absent "${invalid_replay_target_output_path}" "Execution failed for task" \
    "replay no-such-target leaked a Gradle task failure instead of wrapper-level guidance"

set +e
capture_output "${missing_replay_input_output_path}" \
    "${repo_root}/jazzer/bin/replay" protocol-request /tmp/no-such-replay-input.bin
missing_replay_input_exit=$?
set -e

[[ ${missing_replay_input_exit} -eq 2 ]] || die \
    "replay missing input exited ${missing_replay_input_exit}; expected wrapper rejection with code 2"
require_file_contains "${missing_replay_input_output_path}" "Replay input does not exist:" \
    "replay missing input no longer names the missing file"
require_file_contains "${missing_replay_input_output_path}" "Usage: jazzer/bin/replay" \
    "replay missing input no longer prints replay usage guidance"
require_file_absent "${missing_replay_input_output_path}" "Exception in thread" \
    "replay missing input leaked a Java stack trace"
require_file_absent "${missing_replay_input_output_path}" "Execution failed for task" \
    "replay missing input leaked a Gradle task failure instead of wrapper-level guidance"

set +e
capture_output "${invalid_promote_target_output_path}" \
    "${repo_root}/jazzer/bin/promote" no-such-target /tmp/does-not-matter.bin sample-seed
invalid_promote_target_exit=$?
set -e

[[ ${invalid_promote_target_exit} -eq 2 ]] || die \
    "promote no-such-target exited ${invalid_promote_target_exit}; expected wrapper rejection with code 2"
require_file_contains "${invalid_promote_target_output_path}" "Unknown Jazzer target: no-such-target" \
    "promote no-such-target no longer explains the invalid target"
require_file_contains "${invalid_promote_target_output_path}" "Usage: jazzer/bin/promote" \
    "promote no-such-target no longer prints promote usage guidance"
require_file_absent "${invalid_promote_target_output_path}" "Usage: ${repo_root}/jazzer/bin/_run-task" \
    "promote no-such-target fell back to the internal _run-task usage surface"

set +e
capture_output "${missing_promote_input_output_path}" \
    "${repo_root}/jazzer/bin/promote" protocol-request /tmp/no-such-promote-input.bin sample-seed
missing_promote_input_exit=$?
set -e

[[ ${missing_promote_input_exit} -eq 2 ]] || die \
    "promote missing input exited ${missing_promote_input_exit}; expected wrapper rejection with code 2"
require_file_contains "${missing_promote_input_output_path}" "Promotion input does not exist:" \
    "promote missing input no longer names the missing file"
require_file_contains "${missing_promote_input_output_path}" "Usage: jazzer/bin/promote" \
    "promote missing input no longer prints promote usage guidance"
require_file_absent "${missing_promote_input_output_path}" "Exception in thread" \
    "promote missing input leaked a Java stack trace"
require_file_absent "${missing_promote_input_output_path}" "Execution failed for task" \
    "promote missing input leaked a Gradle task failure instead of wrapper-level guidance"

printf 'jazzer-wrapper-ux regression: success\n'
