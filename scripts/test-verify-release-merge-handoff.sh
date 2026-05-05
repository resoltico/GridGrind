#!/usr/bin/env bash
# Exercise the release merge-handoff verifier against disposable Git repositories plus fake Git and
# GitHub CLIs so the post-merge CI wait gate stays regression-tested without depending on network
# or local clone/push transport behavior.

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
    local worktree_dir="${target_dir}/worktree"

    (
        git init "${worktree_dir}" >/dev/null
        cd "${worktree_dir}"
        git config user.name "GridGrind Test"
        git config user.email "gridgrind-test@example.com"
        git checkout -b main >/dev/null 2>&1
        git remote add origin "${target_dir}/origin.git"
        cat > gradle.properties <<EOF
version=${version}
gridgrindDescription=GridGrind test description
EOF
        git add gradle.properties
        git commit -m "Initial merge candidate" >/dev/null
    )

    printf '%s\n' "${worktree_dir}"
}

readonly script_dir="$(resolve_script_dir)"
readonly repo_root="$(cd -P -- "${script_dir}/.." && pwd)"
readonly verify_script="${repo_root}/scripts/verify-release-merge-handoff.sh"
readonly real_git="$(command -v git)"

[[ -x "${verify_script}" ]] || die "missing executable verifier script at ${verify_script}"
[[ -n "${real_git}" ]] || die "failed to resolve git executable"

readonly temp_parent="${repo_root}/tmp/test-verify-release-merge-handoff"
mkdir -p "${temp_parent}"
test_root="${temp_parent}/run.$$"
rm -rf "${test_root}"
mkdir -p "${test_root}"
cleanup() {
    rm -rf "${test_root}"
}
trap cleanup EXIT

readonly fake_bin="${test_root}/bin"
mkdir -p "${fake_bin}"

cat > "${fake_bin}/gh" <<'EOF'
#!/usr/bin/env bash
set -euo pipefail

subcommand="${1:-}"
case "${subcommand}" in
    repo)
        [[ "${2:-}" == "view" ]] || exit 1
        args="$*"
        case "${args}" in
            *"--json nameWithOwner"*".nameWithOwner"*)
                printf '%s\n' "${FAKE_GH_REPO_FULL_NAME:?}"
                ;;
            *"--json defaultBranchRef"*".defaultBranchRef.name"*)
                printf '%s\n' "${FAKE_GH_DEFAULT_BRANCH:?}"
                ;;
            *)
                printf 'unsupported gh repo view invocation: %s\n' "$*" >&2
                exit 1
                ;;
        esac
        ;;
    api)
        path="${2:-}"
        case "${path}" in
            */check-runs?per_page=100)
                if [[ -n "${FAKE_GH_CHECK_RUNS_FILE:-}" ]]; then
                    cat "${FAKE_GH_CHECK_RUNS_FILE}"
                else
                    printf '%s\n' "${FAKE_GH_CHECK_RUNS_TSV:-}"
                fi
                ;;
            *)
                printf 'unsupported gh api invocation: %s\n' "$*" >&2
                exit 1
                ;;
        esac
        ;;
    *)
        printf 'unsupported gh invocation: %s\n' "$*" >&2
        exit 1
        ;;
esac
EOF
chmod +x "${fake_bin}/gh"

cat > "${fake_bin}/git" <<EOF
#!/usr/bin/env bash
set -euo pipefail

real_git="$(printf '%s' "${real_git}")"

if [[ "\${1:-}" == "fetch" && "\${2:-}" == "--no-tags" && "\${3:-}" == "origin" ]]; then
    branch_name="\${4:-}"
    [[ -n "\${FAKE_REMOTE_MAIN_SHA:-}" ]] || {
        printf 'missing FAKE_REMOTE_MAIN_SHA for fake git fetch\n' >&2
        exit 1
    }
    "\${real_git}" update-ref "refs/remotes/origin/\${branch_name}" "\${FAKE_REMOTE_MAIN_SHA}"
    exit 0
fi

exec "\${real_git}" "\$@"
EOF
chmod +x "${fake_bin}/git"

run_verify_expect_success() {
    local worktree_dir=$1
    local remote_head_sha=$2
    shift
    shift
    (
        cd "${worktree_dir}"
        PATH="${fake_bin}:${PATH}" \
            FAKE_GH_REPO_FULL_NAME="example/gridgrind" \
            FAKE_GH_DEFAULT_BRANCH="main" \
            FAKE_REMOTE_MAIN_SHA="${remote_head_sha}" \
            "$@" \
            "${verify_script}" >/dev/null
    )
}

run_verify_expect_failure() {
    local worktree_dir=$1
    local remote_head_sha=$2
    shift
    shift
    (
        cd "${worktree_dir}"
        if PATH="${fake_bin}:${PATH}" \
            FAKE_GH_REPO_FULL_NAME="example/gridgrind" \
            FAKE_GH_DEFAULT_BRANCH="main" \
            FAKE_REMOTE_MAIN_SHA="${remote_head_sha}" \
            "$@" \
            "${verify_script}" >/dev/null 2>&1; then
            die "verifier unexpectedly succeeded"
        fi
    )
}

success_repo="$(create_repo "${test_root}/success" "9.9.9")"
readonly success_sha="$(git -C "${success_repo}" rev-parse HEAD)"
successful_checks="$(printf 'Gate\tcompleted\tsuccess')"
run_verify_expect_success \
    "${success_repo}" \
    "${success_sha}" \
    env \
    FAKE_GH_CHECK_RUNS_TSV="${successful_checks}" \
    GRIDGRIND_RELEASE_CHECK_TIMEOUT_SECONDS=0

delayed_repo="$(create_repo "${test_root}/delayed" "8.8.8")"
readonly delayed_sha="$(git -C "${delayed_repo}" rev-parse HEAD)"
readonly delayed_checks_file="${test_root}/delayed-checks.tsv"
printf 'Gate\tin_progress\t\n' > "${delayed_checks_file}"
(
    sleep 1
    printf '%s\n' "${successful_checks}" > "${delayed_checks_file}"
) &
run_verify_expect_success \
    "${delayed_repo}" \
    "${delayed_sha}" \
    env \
    FAKE_GH_CHECK_RUNS_FILE="${delayed_checks_file}" \
    GRIDGRIND_RELEASE_CHECK_POLL_INTERVAL_SECONDS=1 \
    GRIDGRIND_RELEASE_CHECK_TIMEOUT_SECONDS=5

failure_repo="$(create_repo "${test_root}/failure" "7.7.7")"
readonly failure_sha="$(git -C "${failure_repo}" rev-parse HEAD)"
run_verify_expect_failure \
    "${failure_repo}" \
    "${failure_sha}" \
    env \
    FAKE_GH_CHECK_RUNS_TSV="$(printf 'Gate\tcompleted\tfailure')" \
    GRIDGRIND_RELEASE_CHECK_TIMEOUT_SECONDS=0

stale_repo="$(create_repo "${test_root}/stale" "6.6.6")"
readonly stale_local_sha="$(git -C "${stale_repo}" rev-parse HEAD)"
(
    cd "${stale_repo}"
    cat > gradle.properties <<'EOF'
version=6.6.6
gridgrindDescription=GridGrind updated test description
EOF
    git add gradle.properties
    git commit -m "Advance origin main" >/dev/null
)
readonly stale_remote_sha="$(git -C "${stale_repo}" rev-parse HEAD)"
git -C "${stale_repo}" reset --hard "${stale_local_sha}" >/dev/null
run_verify_expect_failure \
    "${stale_repo}" \
    "${stale_remote_sha}" \
    env \
    FAKE_GH_CHECK_RUNS_TSV="${successful_checks}" \
    GRIDGRIND_RELEASE_CHECK_TIMEOUT_SECONDS=0

printf 'verify-release-merge-handoff regression: success\n'
