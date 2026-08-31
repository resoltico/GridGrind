#!/usr/bin/env bash
# Run the separately scheduled critical mutation-testing gate.

set -euo pipefail

die() { printf 'error: %s\n' "$1" >&2; exit 1; }

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

require_java_26() {
    local version_line javac_version_line
    command -v java >/dev/null 2>&1 || die "java is required"
    command -v javac >/dev/null 2>&1 || die "a full Java 26 JDK is required"
    version_line="$(java --version 2>/dev/null | head -1 || true)"
    [[ "${version_line}" =~ ^[^[:space:]]+[[:space:]]+26([.]|[[:space:]]|$) ]] || die "java must report version 26; found '${version_line:-unavailable}'"
    javac_version_line="$(javac --version 2>/dev/null | head -1 || true)"
    [[ "${javac_version_line}" =~ ^javac[[:space:]]+26([.]|[[:space:]]|$) ]] || die "javac must report version 26; found '${javac_version_line:-unavailable}'"
}

validate_arguments() {
    local argument
    for argument in "$@"; do
        case "${argument}" in
            --dry-run|-m|--stacktrace|--full-stacktrace|-s|-S|--info|--debug|--warn|--quiet|-i|-q|--scan|--profile|--continue|--no-continue|--build-cache|--no-build-cache|--configuration-cache|--no-configuration-cache|--rerun-tasks|--refresh-dependencies|--offline|--warning-mode=*|-D*=*|-P*=*) ;;
            --project-dir|--project-dir=*|-p|-p=*|--build-file|--build-file=*|-b|-b=*|--settings-file|--settings-file=*|-c|-c=*) die "project-location overrides are not supported: ${argument}" ;;
            -*) die "unsupported Gradle option: ${argument}" ;;
            *) die "positional Gradle tasks or values are not supported: ${argument}" ;;
        esac
    done
}

readonly repo_root="$(resolve_script_dir)"
readonly gradlew="${repo_root}/gradlew"
readonly lock_dir="${repo_root}/tmp/repo-verification-lock"
readonly pid_file="${lock_dir}/pid"
lock_scope_name='GridGrind mutation verification command'
lock_scope_advice='wait for the active verification command, then rerun ./check_mutation.sh'
readonly lock_support="${repo_root}/scripts/repo-verification-lock-support.sh"

if [[ "${1:-}" == '-h' || "${1:-}" == '--help' ]]; then
    printf '%s\n' 'Usage: ./check_mutation.sh [supported Gradle options]' '' \
        'Runs the reviewed critical PIT mutationCheck outside the normal ./check.sh gate.' \
        'Reports: contract/build/reports/pitest, engine/build/reports/pitest, cli/build/reports/pitest, and executor/build/reports/pitest'
    exit 0
fi

[[ -x "${gradlew}" ]] || die "missing executable Gradle wrapper at ${gradlew}"
[[ -f "${lock_support}" ]] || die "missing verification lock support at ${lock_support}"
require_java_26
validate_arguments "$@"
# shellcheck source=/dev/null
source "${lock_support}"
trap cleanup_lock EXIT INT TERM
acquire_lock

export GRADLE_USER_HOME="${GRIDGRIND_MUTATION_GRADLE_USER_HOME:-${repo_root}/tmp/mutation-gradle-user-home}"
"${gradlew}" --no-daemon --no-parallel --console=plain mutationCheck "$@"
printf '%s\n' 'Mutation check: success' 'Reports: contract/build/reports/pitest, engine/build/reports/pitest, cli/build/reports/pitest, and executor/build/reports/pitest'
