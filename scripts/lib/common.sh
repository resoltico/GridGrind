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

require_java_26() {
    local tool output version
    for tool in java javac; do
        if ! output="$("${tool}" --version 2>&1)"; then
            printf 'error: %s must resolve to a working Java 26 JDK\n' "${tool}" >&2
            return 1
        fi
        version="${output#* }"
        version="${version%%[[:space:]]*}"
        case "${version}" in
            26|26.*|26+*) ;;
            *) printf 'error: %s must report Java 26; found %s\n' "${tool}" "${output%%$'\n'*}" >&2; return 1 ;;
        esac
    done
}
