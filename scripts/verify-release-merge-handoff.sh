#!/usr/bin/env bash
# Verify that the merged release commit is exactly the current remote default-branch head and wait
# for the release-blocking CI checks on that commit before any release tag is created.

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

trimmed_csv_entries() {
    local csv_value=$1
    while IFS= read -r raw_entry || [[ -n "${raw_entry}" ]]; do
        printf '%s\n' "${raw_entry}" | sed -E 's/^[[:space:]]+//; s/[[:space:]]+$//'
    done < <(printf '%s' "${csv_value}" | tr ',' '\n')
}

format_observed_checks() {
    local output_var_name=$1
    local checks_tsv_path=$2
    local formatted_checks=''
    if [[ ! -s "${checks_tsv_path}" ]]; then
        printf -v "${output_var_name}" '%s' 'none'
        return 0
    fi
    formatted_checks="$(
        awk -F '\t' '
        BEGIN { separator = "" }
        {
            printf "%s%s[%s/%s]", separator, $1, $2, $3
            separator = ", "
        }
        ' "${checks_tsv_path}"
    )"
    printf -v "${output_var_name}" '%s' "${formatted_checks}"
}

blocking_check_state() {
    local output_var_name=$1
    local checks_tsv_path=$2
    local blocking_check_name=$3
    local resolved_state=''
    resolved_state="$(
        awk -F '\t' -v target="${blocking_check_name}" '
        BEGIN {
            has_success = 0
            has_pending = 0
            has_failure = 0
        }
        $1 == target {
            if ($2 == "completed" && $3 == "success") {
                has_success = 1
            } else if ($2 == "completed") {
                has_failure = 1
            } else {
                has_pending = 1
            }
        }
        END {
            if (has_success) {
                print "success"
            } else if (has_pending) {
                print "pending"
            } else if (has_failure) {
                print "failure"
            } else {
                print "missing"
            }
        }
        ' "${checks_tsv_path}"
    )"
    printf -v "${output_var_name}" '%s' "${resolved_state}"
}

require_non_negative_integer() {
    local value=$1
    local name=$2
    [[ "${value}" =~ ^[0-9]+$ ]] || die "${name} must be a non-negative integer, got '${value}'"
}

readonly script_dir="$(resolve_script_dir)"
readonly script_repo_root="$(cd -P -- "${script_dir}/.." && pwd)"
readonly blocking_checks_csv="Gate"
readonly poll_interval_seconds="${GRIDGRIND_RELEASE_CHECK_POLL_INTERVAL_SECONDS:-10}"
readonly timeout_seconds="${GRIDGRIND_RELEASE_CHECK_TIMEOUT_SECONDS:-3600}"

require_non_negative_integer "${poll_interval_seconds}" "GRIDGRIND_RELEASE_CHECK_POLL_INTERVAL_SECONDS"
require_non_negative_integer "${timeout_seconds}" "GRIDGRIND_RELEASE_CHECK_TIMEOUT_SECONDS"

repo_root="${GRIDGRIND_RELEASE_WORKTREE:-$(git rev-parse --show-toplevel 2>/dev/null || true)}"
if [[ -z "${repo_root}" ]]; then
    repo_root="${script_repo_root}"
fi
readonly repo_root
readonly temp_parent="${repo_root}/tmp/verify-release-merge-handoff"

mkdir -p "${temp_parent}"
temp_dir="$(mktemp -d "${temp_parent%/}/run.XXXXXX")"
cleanup() {
    rm -rf "${temp_dir}"
}
trap cleanup EXIT

capture_command_output() {
    local output_path=$1
    shift
    "$@" > "${output_path}"
}

cd "${repo_root}"

readonly target_commit_sha="${1:-$(git rev-parse HEAD)}"
readonly local_head_sha="$(git rev-parse HEAD)"

[[ "${target_commit_sha}" == "${local_head_sha}" ]] || die \
    "target commit ${target_commit_sha} does not match checked-out HEAD ${local_head_sha}"

blocking_check_names=()
while IFS= read -r trimmed_check_name || [[ -n "${trimmed_check_name}" ]]; do
    [[ -n "${trimmed_check_name}" ]] || continue
    blocking_check_names+=("${trimmed_check_name}")
done < <(trimmed_csv_entries "${blocking_checks_csv}")

((${#blocking_check_names[@]} > 0)) || die "no release-blocking merge-handoff checks configured"

readonly repo_full_name_path="${temp_dir}/repo-full-name.txt"
readonly default_branch_path="${temp_dir}/default-branch.txt"
readonly check_runs_tsv_path="${temp_dir}/check-runs.tsv"

capture_command_output \
    "${repo_full_name_path}" \
    gh repo view --json nameWithOwner --jq '.nameWithOwner'
capture_command_output \
    "${default_branch_path}" \
    gh repo view --json defaultBranchRef --jq '.defaultBranchRef.name'

IFS= read -r repo_full_name < "${repo_full_name_path}" || true
IFS= read -r default_branch < "${default_branch_path}" || true
readonly repo_full_name
readonly default_branch

[[ -n "${repo_full_name}" ]] || die "failed to resolve repository name"
[[ -n "${default_branch}" ]] || die "failed to resolve default branch"

git fetch --no-tags origin "${default_branch}" >/dev/null 2>&1 || die \
    "failed to fetch origin/${default_branch}"

readonly remote_default_ref="refs/remotes/origin/${default_branch}"
git show-ref --verify --quiet "${remote_default_ref}" || die \
    "missing ${remote_default_ref} after fetch"

readonly remote_default_sha="$(git rev-parse "${remote_default_ref}")"
[[ "${local_head_sha}" == "${remote_default_sha}" ]] || die \
    "checked-out HEAD ${local_head_sha} does not match origin/${default_branch} ${remote_default_sha}; pull the merged release commit before tagging"

readonly deadline_epoch="$((SECONDS + timeout_seconds))"

while true; do
    capture_command_output \
        "${check_runs_tsv_path}" \
        gh api \
            "/repos/${repo_full_name}/commits/${target_commit_sha}/check-runs?per_page=100" \
            --jq '.check_runs[]? | [.name, .status, .conclusion] | @tsv'

    pending_checks=()
    failed_checks=()
    blocking_state=''
    for blocking_check_name in "${blocking_check_names[@]}"; do
        blocking_check_state blocking_state "${check_runs_tsv_path}" "${blocking_check_name}"
        case "${blocking_state}" in
            success) ;;
            pending|missing)
                pending_checks+=("${blocking_check_name}")
                ;;
            failure)
                failed_checks+=("${blocking_check_name}")
                ;;
            *)
                die "unsupported release-blocking-check state for ${blocking_check_name}"
                ;;
        esac
    done

    observed_checks=''
    format_observed_checks observed_checks "${check_runs_tsv_path}"

    if ((${#failed_checks[@]} > 0)); then
        die \
            "release merge handoff failed for ${target_commit_sha}; release-blocking checks failed (${failed_checks[*]}). Observed check runs: ${observed_checks}"
    fi

    if ((${#pending_checks[@]} == 0)); then
        printf 'Verified release merge handoff at %s on origin/%s with release-blocking checks: %s\n' \
            "${target_commit_sha}" \
            "${default_branch}" \
            "$(IFS=', '; printf '%s' "${blocking_check_names[*]}")"
        exit 0
    fi

    if ((SECONDS >= deadline_epoch)); then
        die \
            "timed out waiting for release merge handoff release-blocking checks (${pending_checks[*]}) on ${target_commit_sha}. Observed check runs: ${observed_checks}"
    fi

    remaining_seconds="$((deadline_epoch - SECONDS))"
    if ((remaining_seconds < 0)); then
        remaining_seconds=0
    fi

    printf 'Waiting for release merge handoff release-blocking checks on %s (%ss remaining). Pending: %s. Observed: %s\n' \
        "${target_commit_sha}" \
        "${remaining_seconds}" \
        "$(IFS=', '; printf '%s' "${pending_checks[*]}")" \
        "${observed_checks}"
    sleep "${poll_interval_seconds}"
done
