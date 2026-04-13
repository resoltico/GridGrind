#!/usr/bin/env bash
# Exercise the public-container verifier against a fake Docker CLI so the release workflow
# contract is tested locally without requiring a real registry push.

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
readonly verify_script="${repo_root}/scripts/verify-container-publication.sh"
readonly expected_description="$(
    awk -F= '
        $1 == "gridgrindDescription" {
            sub(/^[^=]*=/, "", $0)
            print $0
            exit
        }
    ' "${repo_root}/gradle.properties"
)"

[[ -x "${verify_script}" ]] || die "missing executable verifier script at ${verify_script}"
[[ -n "${expected_description}" ]] || die "missing gridgrindDescription in gradle.properties"

test_root="$(mktemp -d)"
cleanup() {
    rm -rf "${test_root}"
}
trap cleanup EXIT

readonly fake_bin="${test_root}/bin"
readonly fake_log="${test_root}/docker.log"
mkdir -p "${fake_bin}"

cat > "${fake_bin}/docker" <<'EOF'
#!/usr/bin/env bash
set -euo pipefail

log_path=${FAKE_DOCKER_LOG:?}
printf '%s\n' "$*" >> "${log_path}"

args=("$@")
if [[ ${#args[@]} -lt 3 || "${args[0]}" != "--config" ]]; then
    printf 'unexpected docker invocation: %s\n' "$*" >&2
    exit 1
fi

command=${args[2]}
case "${command}" in
    pull)
        image_ref=${args[3]:-}
        [[ -n "${image_ref}" ]] || exit 1
        exit 0
        ;;
    run)
        image_ref=${args[4]:-}
        case "${image_ref}" in
            *:latest)
                printf '%s' "${FAKE_DOCKER_LATEST_OUTPUT:?}"
                ;;
            *)
                printf '%s' "${FAKE_DOCKER_VERSION_OUTPUT:?}"
                ;;
        esac
        ;;
    *)
        printf 'unexpected docker subcommand: %s\n' "${command}" >&2
        exit 1
        ;;
esac
EOF
chmod +x "${fake_bin}/docker"

run_verify_expect_success() {
    PATH="${fake_bin}:${PATH}" \
        FAKE_DOCKER_LOG="${fake_log}" \
        FAKE_DOCKER_VERSION_OUTPUT="$1" \
        FAKE_DOCKER_LATEST_OUTPUT="$2" \
        GRIDGRIND_PUBLICATION_VERIFY_RETRIES=1 \
        GRIDGRIND_PUBLICATION_VERIFY_DELAY_SECONDS=0 \
        "${verify_script}" "ghcr.io/example/gridgrind" "9.9.9" >/dev/null
}

run_verify_expect_failure() {
    if PATH="${fake_bin}:${PATH}" \
        FAKE_DOCKER_LOG="${fake_log}" \
        FAKE_DOCKER_VERSION_OUTPUT="$1" \
        FAKE_DOCKER_LATEST_OUTPUT="$2" \
        GRIDGRIND_PUBLICATION_VERIFY_RETRIES=1 \
        GRIDGRIND_PUBLICATION_VERIFY_DELAY_SECONDS=0 \
        "${verify_script}" "ghcr.io/example/gridgrind" "9.9.9" >/dev/null 2>&1; then
        die "verifier unexpectedly succeeded"
    fi
}

expected_header="$(printf 'GridGrind 9.9.9\n%s' "${expected_description}")"

: > "${fake_log}"
run_verify_expect_success "${expected_header}" "${expected_header}"
grep -Fq 'pull ghcr.io/example/gridgrind:9.9.9' "${fake_log}" || die "verifier did not pull the version tag"
grep -Fq 'pull ghcr.io/example/gridgrind:latest' "${fake_log}" || die "verifier did not pull the latest tag"

run_verify_expect_failure "$(printf 'gridgrind 9.9.9')" "$(printf 'gridgrind 9.9.9')"
run_verify_expect_failure "$(printf 'GridGrind 9.9.9\nWrong description')" \
    "$(printf 'GridGrind 9.9.9\nWrong description')"

printf 'verify-container-publication regression: success\n'
