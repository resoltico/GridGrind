#!/usr/bin/env bash
# Reproduce and guard the Jazzer launcher run-lock startup race and stale-lock recovery.

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
readonly lock_support="${repo_root}/jazzer/bin/_run-lock-support"

[[ -f "${lock_support}" ]] || die "missing Jazzer lock support helper"

tmp_dir="$(mktemp -d)"
cleanup() {
    rm -rf "${tmp_dir}"
}
trap cleanup EXIT

lock_dir="${tmp_dir}/run-lock"
pid_file="${lock_dir}/pid"

# shellcheck source=/dev/null
source "${lock_support}"

mkdir -p "${lock_dir}"
printf '999999\n' > "${pid_file}"
acquire_lock
[[ -d "${lock_dir}" ]] || die "stale lock recovery removed the lock directory without reclaiming it"
[[ "$(tr -d '[:space:]' < "${pid_file}")" == "$$" ]] || die "stale lock recovery did not rewrite the pid file to the current process"
cleanup_lock

owner_pid=''
publisher_pid=''
mkdir -p "${lock_dir}"
sleep 2 &
owner_pid=$!
(
    sleep 0.2
    printf '%s\n' "${owner_pid}" > "${pid_file}"
) &
publisher_pid=$!

contender_output=''
set +e
contender_output="$(
    lock_dir="${lock_dir}" pid_file="${pid_file}" bash -lc '
        set -euo pipefail
        source "'"${lock_support}"'"
        acquire_lock
    ' 2>&1
)"
contender_status=$?
set -e

wait "${publisher_pid}"
kill "${owner_pid}" 2>/dev/null || true
wait "${owner_pid}" 2>/dev/null || true

[[ ${contender_status} -ne 0 ]] || die "concurrent contender unexpectedly acquired the Jazzer run lock during startup"
printf '%s' "${contender_output}" | grep -F "already running with PID ${owner_pid}" >/dev/null || die \
    "concurrent contender did not report the active owner after waiting for pid publication"

printf 'jazzer-run-lock regression: success\n'
