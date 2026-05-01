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

status_help="$("${repo_root}/jazzer/bin/status" --help 2>&1)"
require_contains "${status_help}" "Usage: jazzer/bin/status" \
    "status --help no longer prints project-owned wrapper usage"
require_absent "${status_help}" "USAGE: gradlew" \
    "status --help leaked Gradle help instead of wrapper usage"
require_absent "${status_help}" "Welcome to Gradle" \
    "status --help booted Gradle instead of returning wrapper usage"

protocol_help="$("${repo_root}/jazzer/bin/fuzz-protocol-request" --help 2>&1)"
require_contains "${protocol_help}" "Usage: jazzer/bin/fuzz-protocol-request" \
    "fuzz-protocol-request --help no longer prints project-owned wrapper usage"
require_absent "${protocol_help}" "USAGE: gradlew" \
    "fuzz-protocol-request --help leaked Gradle help instead of wrapper usage"
require_absent "${protocol_help}" "Welcome to Gradle" \
    "fuzz-protocol-request --help booted Gradle instead of returning wrapper usage"

fuzz_all_help="$("${repo_root}/jazzer/bin/fuzz-all" --help 2>&1)"
require_contains "${fuzz_all_help}" "Usage: jazzer/bin/fuzz-all" \
    "fuzz-all --help no longer prints project-owned wrapper usage"
require_absent "${fuzz_all_help}" "USAGE: gradlew" \
    "fuzz-all --help leaked Gradle help instead of wrapper usage"

replay_help="$("${repo_root}/jazzer/bin/replay" --help 2>&1)"
require_contains "${replay_help}" "Usage: jazzer/bin/replay" \
    "replay --help no longer prints project-owned wrapper usage"
require_contains "${replay_help}" "Valid targets:" \
    "replay --help no longer lists valid replay targets"

promote_help="$("${repo_root}/jazzer/bin/promote" --help 2>&1)"
require_contains "${promote_help}" "Usage: jazzer/bin/promote" \
    "promote --help no longer prints project-owned wrapper usage"
require_contains "${promote_help}" "Valid targets:" \
    "promote --help no longer lists valid promotion targets"

set +e
invalid_report_output="$("${repo_root}/jazzer/bin/report" no-such-target --console=plain 2>&1)"
invalid_report_exit=$?
set -e

[[ ${invalid_report_exit} -eq 2 ]] || die \
    "report no-such-target exited ${invalid_report_exit}; expected the wrapper to reject it with code 2"
require_contains "${invalid_report_output}" "Unknown Jazzer target: no-such-target" \
    "report no-such-target no longer explains the invalid target"
require_contains "${invalid_report_output}" "Usage: jazzer/bin/report" \
    "report no-such-target no longer prints wrapper usage guidance"
require_absent "${invalid_report_output}" "Exception in thread" \
    "report no-such-target leaked a Java stack trace"
require_absent "${invalid_report_output}" "Execution failed for task" \
    "report no-such-target leaked a Gradle task failure instead of a wrapper-level error"
require_absent "${invalid_report_output}" "Welcome to Gradle" \
    "report no-such-target booted Gradle instead of rejecting the target up front"

set +e
invalid_replay_target_output="$("${repo_root}/jazzer/bin/replay" no-such-target /tmp/does-not-matter.bin 2>&1)"
invalid_replay_target_exit=$?
set -e

[[ ${invalid_replay_target_exit} -eq 2 ]] || die \
    "replay no-such-target exited ${invalid_replay_target_exit}; expected wrapper rejection with code 2"
require_contains "${invalid_replay_target_output}" "Unknown Jazzer target: no-such-target" \
    "replay no-such-target no longer explains the invalid target"
require_contains "${invalid_replay_target_output}" "Usage: jazzer/bin/replay" \
    "replay no-such-target no longer prints replay usage guidance"
require_absent "${invalid_replay_target_output}" "Usage: ${repo_root}/jazzer/bin/_run-task" \
    "replay no-such-target fell back to the internal _run-task usage surface"
require_absent "${invalid_replay_target_output}" "Execution failed for task" \
    "replay no-such-target leaked a Gradle task failure instead of wrapper-level guidance"

set +e
missing_replay_input_output="$("${repo_root}/jazzer/bin/replay" protocol-request /tmp/no-such-replay-input.bin 2>&1)"
missing_replay_input_exit=$?
set -e

[[ ${missing_replay_input_exit} -eq 2 ]] || die \
    "replay missing input exited ${missing_replay_input_exit}; expected wrapper rejection with code 2"
require_contains "${missing_replay_input_output}" "Replay input does not exist:" \
    "replay missing input no longer names the missing file"
require_contains "${missing_replay_input_output}" "Usage: jazzer/bin/replay" \
    "replay missing input no longer prints replay usage guidance"
require_absent "${missing_replay_input_output}" "Exception in thread" \
    "replay missing input leaked a Java stack trace"
require_absent "${missing_replay_input_output}" "Execution failed for task" \
    "replay missing input leaked a Gradle task failure instead of wrapper-level guidance"

set +e
invalid_promote_target_output="$("${repo_root}/jazzer/bin/promote" no-such-target /tmp/does-not-matter.bin sample-seed 2>&1)"
invalid_promote_target_exit=$?
set -e

[[ ${invalid_promote_target_exit} -eq 2 ]] || die \
    "promote no-such-target exited ${invalid_promote_target_exit}; expected wrapper rejection with code 2"
require_contains "${invalid_promote_target_output}" "Unknown Jazzer target: no-such-target" \
    "promote no-such-target no longer explains the invalid target"
require_contains "${invalid_promote_target_output}" "Usage: jazzer/bin/promote" \
    "promote no-such-target no longer prints promote usage guidance"
require_absent "${invalid_promote_target_output}" "Usage: ${repo_root}/jazzer/bin/_run-task" \
    "promote no-such-target fell back to the internal _run-task usage surface"

set +e
missing_promote_input_output="$("${repo_root}/jazzer/bin/promote" protocol-request /tmp/no-such-promote-input.bin sample-seed 2>&1)"
missing_promote_input_exit=$?
set -e

[[ ${missing_promote_input_exit} -eq 2 ]] || die \
    "promote missing input exited ${missing_promote_input_exit}; expected wrapper rejection with code 2"
require_contains "${missing_promote_input_output}" "Promotion input does not exist:" \
    "promote missing input no longer names the missing file"
require_contains "${missing_promote_input_output}" "Usage: jazzer/bin/promote" \
    "promote missing input no longer prints promote usage guidance"
require_absent "${missing_promote_input_output}" "Exception in thread" \
    "promote missing input leaked a Java stack trace"
require_absent "${missing_promote_input_output}" "Execution failed for task" \
    "promote missing input leaked a Gradle task failure instead of wrapper-level guidance"

printf 'jazzer-wrapper-ux regression: success\n'
