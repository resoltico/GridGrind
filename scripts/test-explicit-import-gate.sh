#!/usr/bin/env bash
# Prove the explicit-import verification gate fails on a real src/main wildcard import.

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
readonly sentinel_path="${repo_root}/cli/src/main/java/dev/erst/gridgrind/cli/WildcardImportGateSentinel.java"
readonly sentinel_relative_path='cli/src/main/java/dev/erst/gridgrind/cli/WildcardImportGateSentinel.java'
output_file=''
project_cache_dir=''

prepare_project_cache_dir() {
    local cache_root=$1
    if [[ -z "${cache_root}" ]]; then
        return 0
    fi
    mkdir -p "${cache_root}"
    mktemp -d "${cache_root%/}/explicit-import-gate.XXXXXX"
}

project_cache_dir="$(prepare_project_cache_dir "${GRIDGRIND_PROJECT_CACHE_DIR:-}")"

cleanup() {
    rm -f "${sentinel_path}"
    if [[ -n "${output_file}" ]]; then
        rm -f "${output_file}"
    fi
    if [[ -n "${project_cache_dir}" ]]; then
        rm -rf "${project_cache_dir}"
    fi
}

trap cleanup EXIT

[[ ! -e "${sentinel_path}" ]] || die "sentinel path already exists: ${sentinel_path}"

cat >"${sentinel_path}" <<'EOF'
package dev.erst.gridgrind.cli;

import java.util.*;

final class WildcardImportGateSentinel {}
EOF

run_verify_explicit_imports() {
    output_file="$(mktemp)"
    local gradle_pid=''
    local heartbeat_pid=''
    local status=0
    local -a gradle_command=(
        "${repo_root}/gradlew"
        --console=plain
        --no-daemon
    )

    if [[ -n "${project_cache_dir}" ]]; then
        gradle_command+=(--project-cache-dir "${project_cache_dir}")
    fi
    gradle_command+=(verifyExplicitImports)

    "${gradle_command[@]}" >"${output_file}" 2>&1 &
    gradle_pid=$!
    (
        while kill -0 "${gradle_pid}" 2>/dev/null; do
            printf 'explicit-import gate regression: waiting for verifyExplicitImports...\n'
            sleep 15
        done
    ) &
    heartbeat_pid=$!

    set +e
    wait "${gradle_pid}"
    status=$?

    kill "${heartbeat_pid}" 2>/dev/null || true
    wait "${heartbeat_pid}" 2>/dev/null || true
    return "${status}"
}

set +e
run_verify_explicit_imports
status=$?
set -e

[[ ${status} -ne 0 ]] || die 'verifyExplicitImports unexpectedly accepted a wildcard import in src/main'
grep -F "${sentinel_relative_path}:3: import java.util.*;" "${output_file}" >/dev/null \
    || die 'verifyExplicitImports did not report the injected wildcard import precisely'

printf 'explicit-import gate regression: success\n'
