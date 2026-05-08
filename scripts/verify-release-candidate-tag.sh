#!/usr/bin/env bash
# Verify that a tag-targeted publication candidate is safe to publish: the checked-out commit must
# match the remote tag, the tag version must match gradle.properties, the tag commit must still be
# reachable from the default branch, and the release-blocking CI checks must already be green on that
# exact commit.

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
readonly script_repo_root="$(cd -P -- "${script_dir}/.." && pwd)"
readonly tag_name="${1:-${RELEASE_TAG:-${GITHUB_REF_NAME:-}}}"
readonly blocking_checks_csv="Gate"

[[ -n "${tag_name}" ]] || die "release tag is required"
[[ "${tag_name}" == v* ]] || die "release tag must start with v"

repo_root="${GRIDGRIND_RELEASE_WORKTREE:-$(git rev-parse --show-toplevel 2>/dev/null || true)}"
if [[ -z "${repo_root}" ]]; then
    repo_root="${script_repo_root}"
fi
readonly repo_root
readonly temp_parent="${repo_root}/tmp/verify-release-candidate-tag"

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

cd "${repo_root}"

readonly expected_version="${tag_name#v}"
readonly gradle_version="$(
    awk -F= '
        $1 == "version" {
            print $2
            exit
        }
    ' "${repo_root}/gradle.properties"
)"

[[ -n "${gradle_version}" ]] || die "missing version in gradle.properties"
[[ "${expected_version}" == "${gradle_version}" ]] || die \
    "tag version ${expected_version} does not match gradle.properties version ${gradle_version}"

readonly repo_full_name_path="${temp_dir}/repo-full-name.txt"
readonly default_branch_path="${temp_dir}/default-branch.txt"
readonly remote_tag_object_type_path="${temp_dir}/remote-tag-object-type.txt"
readonly remote_tag_object_sha_path="${temp_dir}/remote-tag-object-sha.txt"
readonly tag_commit_sha_path="${temp_dir}/tag-commit-sha.txt"
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
readonly local_commit_sha="$(git rev-parse HEAD)"

[[ -n "${repo_full_name}" ]] || die "failed to resolve repository name"
[[ -n "${default_branch}" ]] || die "failed to resolve default branch"

readonly tag_ref_api="/repos/${repo_full_name}/git/ref/tags/${tag_name}"
capture_command_output "${remote_tag_object_type_path}" gh api "${tag_ref_api}" --jq '.object.type'
capture_command_output "${remote_tag_object_sha_path}" gh api "${tag_ref_api}" --jq '.object.sha'
IFS= read -r remote_tag_object_type < "${remote_tag_object_type_path}" || true
IFS= read -r remote_tag_object_sha < "${remote_tag_object_sha_path}" || true
readonly remote_tag_object_type
readonly remote_tag_object_sha

case "${remote_tag_object_type}" in
    commit)
        readonly tag_commit_sha="${remote_tag_object_sha}"
        ;;
    tag)
        capture_command_output \
            "${tag_commit_sha_path}" \
            gh api "/repos/${repo_full_name}/git/tags/${remote_tag_object_sha}" --jq '.object.sha'
        IFS= read -r tag_commit_sha < "${tag_commit_sha_path}" || true
        readonly tag_commit_sha
        ;;
    *)
        die "unsupported remote tag object type '${remote_tag_object_type}' for ${tag_name}"
        ;;
esac

[[ "${local_commit_sha}" == "${tag_commit_sha}" ]] || die \
    "checked-out commit ${local_commit_sha} does not match remote tag ${tag_name} commit ${tag_commit_sha}"

default_branch_ref="refs/remotes/origin/${default_branch}"
if ! git show-ref --verify --quiet "${default_branch_ref}"; then
    git fetch --no-tags origin \
        "${default_branch}:${default_branch_ref}" >/dev/null 2>&1 || die \
        "failed to fetch origin/${default_branch}"
fi

git show-ref --verify --quiet "${default_branch_ref}" || die \
    "missing ${default_branch_ref} after fetch"

git merge-base --is-ancestor "${tag_commit_sha}" "${default_branch_ref}" || die \
    "tag ${tag_name} commit ${tag_commit_sha} is not reachable from origin/${default_branch}"

capture_command_output \
    "${check_runs_tsv_path}" \
    gh api \
        "/repos/${repo_full_name}/commits/${tag_commit_sha}/check-runs?per_page=100" \
        --jq '.check_runs[] | [.name, .status, .conclusion] | @tsv'

blocking_check_names=()
while IFS= read -r raw_check_name || [[ -n "${raw_check_name}" ]]; do
    trimmed_check_name="$(printf '%s' "${raw_check_name}" | sed -E 's/^[[:space:]]+//; s/[[:space:]]+$//')"
    [[ -n "${trimmed_check_name}" ]] || continue
    blocking_check_names+=("${trimmed_check_name}")
done < <(printf '%s' "${blocking_checks_csv}" | tr ',' '\n')

((${#blocking_check_names[@]} > 0)) || die "no release-blocking publication checks configured"

missing_checks=()
for blocking_check_name in "${blocking_check_names[@]}"; do
    if ! awk -F '\t' -v target="${blocking_check_name}" '
        $1 == target && $2 == "completed" && $3 == "success" { found = 1 }
        END { exit found ? 0 : 1 }
    ' "${check_runs_tsv_path}"; then
        missing_checks+=("${blocking_check_name}")
    fi
done

if ((${#missing_checks[@]} > 0)); then
    observed_checks=''
    format_observed_checks observed_checks "${check_runs_tsv_path}"
    die \
        "tag ${tag_name} commit ${tag_commit_sha} is missing successful release-blocking checks (${missing_checks[*]}). Observed check runs: ${observed_checks}"
fi

printf 'Verified release candidate %s at %s on origin/%s with release-blocking checks: %s\n' \
    "${tag_name}" \
    "${tag_commit_sha}" \
    "${default_branch}" \
    "$(IFS=', '; printf '%s' "${blocking_check_names[*]}")"
