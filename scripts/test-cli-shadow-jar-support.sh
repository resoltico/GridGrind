#!/usr/bin/env bash
# Prove a stale JAR never masks a failed packaged-artifact rebuild.

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
        [[ "${source_path}" == /* ]] || source_path="${source_dir}/${source_path}"
    done
    cd -P -- "$(dirname -- "${source_path}")" && pwd
}

readonly script_dir="$(resolve_script_dir)"
readonly repo_root="$(cd -P -- "${script_dir}/.." && pwd)"
readonly helper_path="${repo_root}/scripts/lib/cli-shadow-jar-support.sh"
readonly test_root="$(mktemp -d "${repo_root}/tmp/test-cli-shadow-jar-support.XXXXXX")"

cleanup() {
    rm -rf "${test_root}"
}
trap cleanup EXIT

mkdir -p "${test_root}/cli/build/libs"
printf 'stale jar\n' > "${test_root}/cli/build/libs/gridgrind.jar"
cat > "${test_root}/gradlew" <<'EOF'
#!/usr/bin/env bash
printf 'forced shadow-jar failure\n' >&2
exit 42
EOF
chmod +x "${test_root}/gradlew"

# shellcheck source=/dev/null
source "${helper_path}"

if output="$(ensure_cli_shadow_jar "${test_root}" 2>&1)"; then
    die "shadow-jar helper accepted a stale artifact after a failed rebuild: ${output}"
fi
grep -Fq 'forced shadow-jar failure' <<<"${output}" || die \
    "shadow-jar helper did not preserve the failed Gradle invocation output"
grep -Fq 'failed to build CLI shadow jar' <<<"${output}" || die \
    "shadow-jar helper did not report its failed rebuild"

printf 'cli shadow-jar support regression: success\n'
