#!/usr/bin/env bash
# Run the canonical Stage 2 quality gates: repo hygiene verification then check coverage.

set -euo pipefail

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
readonly gradlew="${repo_root}/gradlew"
readonly repo_hygiene_verifier="${repo_root}/scripts/verify-repo-hygiene.sh"

print_usage() {
    printf '%s\n' \
        'Usage: ./scripts/run-quality-gates.sh [supported Gradle options]' \
        '' \
        'Runs the canonical Stage 2 verification surface:' \
        '  1. ./scripts/verify-repo-hygiene.sh' \
        '  2. ./gradlew check coverage' \
        '' \
        'Any remaining arguments are forwarded to the Gradle invocation.' \
        'Use ./check.sh for the full six-stage repository gate.'
}

for argument in "$@"; do
    case "${argument}" in
        -h|--help)
            print_usage
            exit 0
            ;;
    esac
done

[[ -x "${gradlew}" ]] || {
    printf 'error: missing executable Gradle wrapper at %s\n' "${gradlew}" >&2
    exit 1
}
[[ -x "${repo_hygiene_verifier}" ]] || {
    printf 'error: missing executable repo hygiene verifier at %s\n' "${repo_hygiene_verifier}" >&2
    exit 1
}

"${repo_hygiene_verifier}"
"${gradlew}" check coverage "$@"
