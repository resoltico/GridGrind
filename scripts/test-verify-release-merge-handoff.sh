#!/usr/bin/env bash
# Exercise the release merge-handoff verifier against disposable Git repositories and a fake
# GitHub CLI so the post-merge CI wait gate stays regression-tested without touching the network.

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
    local worktree_dir="${target_dir}/worktree"

    git init --bare "${origin_dir}" >/dev/null
    git clone "${origin_dir}" "${worktree_dir}" >/dev/null 2>&1

    (
        cd "${worktree_dir}"
        git config user.name "GridGrind Test"
        git config user.email "gridgrind-test@example.com"
        git checkout -b main >/dev/null 2>&1
        cat > gradle.properties <<EOF
version=${version}
gridgrindDescription=GridGrind test description
EOF
        git add gradle.properties
        git commit -m "Initial merge candidate" >/dev/null
        git push -u origin main >/dev/null 2>&1
    )

    printf '%s\n' "${worktree_dir}"
}

readonly script_dir="$(resolve_script_dir)"
readonly repo_root="$(cd -P -- "${script_dir}/.." && pwd)"
readonly verify_script="${repo_root}/scripts/verify-release-merge-handoff.sh"

[[ -x "${verify_script}" ]] || die "missing executable verifier script at ${verify_script}"

test_root="$(mktemp -d)"
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

run_verify_expect_success() {
    local worktree_dir=$1
    shift
    (
        cd "${worktree_dir}"
        PATH="${fake_bin}:${PATH}" \
            FAKE_GH_REPO_FULL_NAME="example/gridgrind" \
            FAKE_GH_DEFAULT_BRANCH="main" \
            "$@" \
            "${verify_script}" >/dev/null
    )
}

run_verify_expect_failure() {
    local worktree_dir=$1
    shift
    (
        cd "${worktree_dir}"
        if PATH="${fake_bin}:${PATH}" \
            FAKE_GH_REPO_FULL_NAME="example/gridgrind" \
            FAKE_GH_DEFAULT_BRANCH="main" \
            "$@" \
            "${verify_script}" >/dev/null 2>&1; then
            die "verifier unexpectedly succeeded"
        fi
    )
}

success_repo="$(create_repo "${test_root}/success" "9.9.9")"
successful_checks="$(printf 'Check\tcompleted\tsuccess\nDocker smoke\tcompleted\tsuccess')"
run_verify_expect_success \
    "${success_repo}" \
    env \
    FAKE_GH_CHECK_RUNS_TSV="${successful_checks}" \
    GRIDGRIND_RELEASE_CHECK_TIMEOUT_SECONDS=0

delayed_repo="$(create_repo "${test_root}/delayed" "8.8.8")"
readonly delayed_checks_file="${test_root}/delayed-checks.tsv"
printf 'Check\tin_progress\t\nDocker smoke\tqueued\t\n' > "${delayed_checks_file}"
(
    sleep 1
    printf '%s\n' "${successful_checks}" > "${delayed_checks_file}"
) &
run_verify_expect_success \
    "${delayed_repo}" \
    env \
    FAKE_GH_CHECK_RUNS_FILE="${delayed_checks_file}" \
    GRIDGRIND_RELEASE_CHECK_POLL_INTERVAL_SECONDS=1 \
    GRIDGRIND_RELEASE_CHECK_TIMEOUT_SECONDS=5

failure_repo="$(create_repo "${test_root}/failure" "7.7.7")"
run_verify_expect_failure \
    "${failure_repo}" \
    env \
    FAKE_GH_CHECK_RUNS_TSV="$(printf 'Check\tcompleted\tsuccess\nDocker smoke\tcompleted\tfailure')" \
    GRIDGRIND_RELEASE_CHECK_TIMEOUT_SECONDS=0

stale_repo="$(create_repo "${test_root}/stale" "6.6.6")"
readonly stale_origin_dir="${test_root}/stale/origin.git"
readonly stale_updater_dir="${test_root}/stale/updater"
git clone "${stale_origin_dir}" "${stale_updater_dir}" >/dev/null 2>&1
(
    cd "${stale_updater_dir}"
    git config user.name "GridGrind Test"
    git config user.email "gridgrind-test@example.com"
    git checkout main >/dev/null 2>&1
    cat > gradle.properties <<'EOF'
version=6.6.6
gridgrindDescription=GridGrind updated test description
EOF
    git add gradle.properties
    git commit -m "Advance origin main" >/dev/null
    git push origin main >/dev/null 2>&1
)
run_verify_expect_failure \
    "${stale_repo}" \
    env \
    FAKE_GH_CHECK_RUNS_TSV="${successful_checks}" \
    GRIDGRIND_RELEASE_CHECK_TIMEOUT_SECONDS=0

printf 'verify-release-merge-handoff regression: success\n'
