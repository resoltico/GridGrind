#!/usr/bin/env bash
# Shared repository location and shell-JDK precondition.

resolve_script_dir() {
    local source_path="$1" source_dir
    while [[ -h "${source_path}" ]]; do
        source_dir="$(cd -P -- "$(dirname -- "${source_path}")" && pwd)"
        source_path="$(readlink "${source_path}")"
        [[ "${source_path}" = /* ]] || source_path="${source_dir}/${source_path}"
    done
    cd -P -- "$(dirname -- "${source_path}")" && pwd
}

GRIDGRIND_ROOT="$(cd "$(resolve_script_dir "${BASH_SOURCE[0]}")/../.." && pwd)"
readonly GRIDGRIND_ROOT

require_java() {
    local tool output version required_version key value
    required_version=""
    while IFS='=' read -r key value; do
        if [[ "${key}" == gridgrindJavaVersion ]]; then required_version="${value}"; fi
    done < "${GRIDGRIND_ROOT}/gradle.properties"
    if [[ ! "${required_version}" =~ ^[0-9]+$ ]]; then
        printf 'error: gradle.properties must declare gridgrindJavaVersion as an integer\n' >&2
        return 1
    fi
    for tool in java javac; do
        if ! output="$("${tool}" --version 2>&1)"; then
            printf 'error: %s must resolve to a working Java %s JDK\n' "${tool}" "${required_version}" >&2
            return 1
        fi
        version="${output#* }"
        version="${version%%[[:space:]]*}"
        case "${version}" in
            "${required_version}"|"${required_version}".*|"${required_version}"+*) ;;
            *) printf 'error: %s must report Java %s; found %s\n' "${tool}" "${required_version}" "${output%%$'\n'*}" >&2; return 1 ;;
        esac
    done
}
