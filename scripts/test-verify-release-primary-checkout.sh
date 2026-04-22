#!/usr/bin/env bash
# Exercise the primary-checkout release-closeout verifier against disposable repositories so the
# post-release reconciliation contract cannot drift back into prose-only guidance.

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

create_repo() {
    local target_dir=$1
    local version=$2
    local origin_dir="${target_dir}/origin.git"
    local primary_dir="${target_dir}/primary"

    git init --bare "${origin_dir}" >/dev/null
    git clone "${origin_dir}" "${primary_dir}" >/dev/null 2>&1
    (
        cd "${primary_dir}"
        git config user.name "GridGrind Test"
        git config user.email "gridgrind-test@example.com"
        git checkout -b main >/dev/null 2>&1
        cat > gradle.properties <<EOF
version=${version}
gridgrindDescription=GridGrind test description
EOF
        cat > CHANGELOG.md <<EOF
# Changelog

## [Unreleased]

## [${version}] - 2026-04-22

### Fixed

- Initial release state.
EOF
        git add gradle.properties CHANGELOG.md
        git commit -m "Release ${version}" >/dev/null
        git push -u origin main >/dev/null 2>&1
    )
    git -C "${origin_dir}" symbolic-ref HEAD refs/heads/main
    printf '%s\n' "${primary_dir}"
}

run_verify_expect_success() {
    local primary_dir=$1
    local expected_version=$2
    shift 2
    (
        "$@" "${verify_script}" "${primary_dir}" "${expected_version}" >/dev/null
    )
}

run_verify_expect_failure() {
    local primary_dir=$1
    local expected_version=$2
    shift 2
    (
        if "$@" "${verify_script}" "${primary_dir}" "${expected_version}" >/dev/null 2>&1; then
            die "verifier unexpectedly succeeded"
        fi
    )
}

readonly script_dir="$(resolve_script_dir)"
readonly repo_root="$(cd -P -- "${script_dir}/.." && pwd)"
readonly verify_script="${repo_root}/scripts/verify-release-primary-checkout.sh"

[[ -x "${verify_script}" ]] || die "missing executable verifier script at ${verify_script}"

readonly temp_parent="${repo_root}/tmp/test-verify-release-primary-checkout"
mkdir -p "${temp_parent}"
test_root="${temp_parent}/run.$$"
rm -rf "${test_root}"
mkdir -p "${test_root}"
cleanup() {
    rm -rf "${test_root}"
}
trap cleanup EXIT

success_repo="$(create_repo "${test_root}/success" "9.9.9")"
mkdir -p "${success_repo}/tmp/release scratch" "${success_repo}/generated"
printf 'scratch\n' > "${success_repo}/tmp/release scratch/log.txt"
printf 'scratch\n' > "${success_repo}/generated/example.txt"
run_verify_expect_success "${success_repo}" "9.9.9" env

unexpected_untracked_repo="$(create_repo "${test_root}/unexpected-untracked" "8.8.8")"
printf 'oops\n' > "${unexpected_untracked_repo}/unexpected.txt"
run_verify_expect_failure "${unexpected_untracked_repo}" "8.8.8" env

dirty_repo="$(create_repo "${test_root}/dirty" "7.7.7")"
printf '\nreleaseNote=dirty\n' >> "${dirty_repo}/gradle.properties"
run_verify_expect_failure "${dirty_repo}" "7.7.7" env

wrong_branch_repo="$(create_repo "${test_root}/wrong-branch" "6.6.6")"
git -C "${wrong_branch_repo}" checkout -b feature/reconcile >/dev/null 2>&1
run_verify_expect_failure "${wrong_branch_repo}" "6.6.6" env

stale_repo="$(create_repo "${test_root}/stale" "5.5.5")"
peer_repo="${test_root}/stale-peer"
git clone "${test_root}/stale/origin.git" "${peer_repo}" >/dev/null 2>&1
(
    cd "${peer_repo}"
    git config user.name "GridGrind Test"
    git config user.email "gridgrind-test@example.com"
    git checkout main >/dev/null 2>&1
    printf '\nreleaseNote=peer\n' >> gradle.properties
    git add gradle.properties
    git commit -m "Advance origin" >/dev/null
    git push origin main >/dev/null 2>&1
)
run_verify_expect_failure "${stale_repo}" "5.5.5" env

version_repo="$(create_repo "${test_root}/version" "4.4.4")"
run_verify_expect_failure "${version_repo}" "4.4.5" env

printf 'verify-release-primary-checkout regression: success\n'
