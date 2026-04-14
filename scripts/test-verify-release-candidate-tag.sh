#!/usr/bin/env bash
# Exercise the release-candidate verifier against disposable Git repositories and a fake GitHub
# CLI so the publication gate stays regression-tested without touching the network.

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
        git commit -m "Initial release candidate" >/dev/null
        git push -u origin main >/dev/null 2>&1
    )

    printf '%s\n' "${worktree_dir}"
}

readonly script_dir="$(resolve_script_dir)"
readonly repo_root="$(cd -P -- "${script_dir}/.." && pwd)"
readonly verify_script="${repo_root}/scripts/verify-release-candidate-tag.sh"

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
            */git/ref/tags/*)
                case "$*" in
                    *".object.type"*)
                        printf '%s\n' "${FAKE_GH_TAG_OBJECT_TYPE:?}"
                        ;;
                    *".object.sha"*)
                        printf '%s\n' "${FAKE_GH_TAG_OBJECT_SHA:?}"
                        ;;
                    *)
                        printf 'unsupported gh api tag-ref invocation: %s\n' "$*" >&2
                        exit 1
                        ;;
                esac
                ;;
            */git/tags/*)
                case "$*" in
                    *".object.sha"*)
                        printf '%s\n' "${FAKE_GH_TAG_COMMIT_SHA:?}"
                        ;;
                    *)
                        printf 'unsupported gh api annotated-tag invocation: %s\n' "$*" >&2
                        exit 1
                        ;;
                esac
                ;;
            */check-runs?per_page=100)
                printf '%s\n' "${FAKE_GH_CHECK_RUNS_TSV:-}"
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
    local tag_name=$2
    local tag_sha=$3
    shift 3
    (
        cd "${worktree_dir}"
        PATH="${fake_bin}:${PATH}" \
            FAKE_GH_REPO_FULL_NAME="example/gridgrind" \
            FAKE_GH_DEFAULT_BRANCH="main" \
            FAKE_GH_TAG_OBJECT_TYPE="commit" \
            FAKE_GH_TAG_OBJECT_SHA="${tag_sha}" \
            FAKE_GH_TAG_COMMIT_SHA="${tag_sha}" \
            FAKE_GH_CHECK_RUNS_TSV="$1" \
            "${verify_script}" "${tag_name}" >/dev/null
    )
}

run_verify_expect_failure() {
    local worktree_dir=$1
    local tag_name=$2
    local tag_sha=$3
    local check_runs_tsv=$4
    (
        cd "${worktree_dir}"
        if PATH="${fake_bin}:${PATH}" \
            FAKE_GH_REPO_FULL_NAME="example/gridgrind" \
            FAKE_GH_DEFAULT_BRANCH="main" \
            FAKE_GH_TAG_OBJECT_TYPE="commit" \
            FAKE_GH_TAG_OBJECT_SHA="${tag_sha}" \
            FAKE_GH_TAG_COMMIT_SHA="${tag_sha}" \
            FAKE_GH_CHECK_RUNS_TSV="${check_runs_tsv}" \
            "${verify_script}" "${tag_name}" >/dev/null 2>&1; then
            die "verifier unexpectedly succeeded"
        fi
    )
}

success_repo="$(create_repo "${test_root}/success" "9.9.9")"
success_sha="$(git -C "${success_repo}" rev-parse HEAD)"
successful_checks="$(printf 'Check\tcompleted\tsuccess\nDocker smoke\tcompleted\tsuccess')"
run_verify_expect_success "${success_repo}" "v9.9.9" "${success_sha}" "${successful_checks}"

(
    cd "${success_repo}"
    cat > gradle.properties <<'EOF'
version=9.9.8
gridgrindDescription=GridGrind test description
EOF
)
run_verify_expect_failure "${success_repo}" "v9.9.9" "${success_sha}" "${successful_checks}"

branch_repo="$(create_repo "${test_root}/branch" "8.8.8")"
(
    cd "${branch_repo}"
    git checkout -b release-detour >/dev/null 2>&1
    cat > gradle.properties <<'EOF'
version=8.8.8
gridgrindDescription=GridGrind branch-detour description
EOF
    git add gradle.properties
    git commit -m "Release candidate off main" >/dev/null
)
branch_sha="$(git -C "${branch_repo}" rev-parse HEAD)"
run_verify_expect_failure "${branch_repo}" "v8.8.8" "${branch_sha}" "${successful_checks}"

checks_repo="$(create_repo "${test_root}/checks" "7.7.7")"
checks_sha="$(git -C "${checks_repo}" rev-parse HEAD)"
run_verify_expect_failure \
    "${checks_repo}" \
    "v7.7.7" \
    "${checks_sha}" \
    "$(printf 'Check\tcompleted\tsuccess\nDocker smoke\tcompleted\tfailure')"

printf 'verify-release-candidate-tag regression: success\n'
