#!/usr/bin/env bash
# Keep the shared repo-verification lock wired into every top-level verification entrypoint.

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
readonly lock_support="${repo_root}/scripts/repo-verification-lock-support.sh"

[[ -f "${lock_support}" ]] || die "missing repo verification lock helper"
grep -Fq 'repo-verification-lock-support.sh' "${repo_root}/check.sh" || die \
    "check.sh no longer sources the repo verification lock helper"
grep -Fq 'repo-verification-lock-support.sh' "${repo_root}/scripts/docker-smoke.sh" || die \
    "docker-smoke.sh no longer sources the repo verification lock helper"
grep -Fq 'repo-verification-lock-support.sh' "${repo_root}/scripts/validate-devcontainer.sh" || die \
    "validate-devcontainer.sh no longer sources the repo verification lock helper"
grep -Fq 'repo-verification-lock-support.sh' "${repo_root}/jazzer/bin/_run-lock-support" || die \
    "Jazzer lock support no longer routes through the repo verification lock helper"

tmp_dir="$(mktemp -d)"
cleanup() {
    rm -rf "${tmp_dir}"
}
trap cleanup EXIT

lock_dir="${tmp_dir}/repo-lock"
pid_file="${lock_dir}/pid"

# shellcheck source=/dev/null
source "${lock_support}"

mkdir -p "${lock_dir}"
printf '999999\n' > "${pid_file}"
acquire_lock
[[ -d "${lock_dir}" ]] || die "stale lock recovery removed the repo verification lock directory"
[[ "$(tr -d '[:space:]' < "${pid_file}")" == "$$" ]] || die \
    "stale lock recovery did not rewrite the repo verification lock pid file"
cleanup_lock

mkdir -p "${lock_dir}"
printf '%s\n' "$$" > "${pid_file}"
acquire_lock
[[ "${lock_is_reentrant}" == true ]] || die "self-reentrant repo verification lock acquisition was not allowed"
cleanup_lock
rm -rf "${lock_dir}"

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

[[ ${contender_status} -ne 0 ]] || die "concurrent contender unexpectedly acquired the repo verification lock"
printf '%s' "${contender_output}" | grep -F "already running with PID ${owner_pid}" >/dev/null || die \
    "concurrent contender did not report the active repo verification owner"
printf '%s' "${contender_output}" | grep -F 'GridGrind verification command' >/dev/null || die \
    "concurrent contender did not report the repo verification lock scope"

printf 'repo-verification-lock regression: success\n'
