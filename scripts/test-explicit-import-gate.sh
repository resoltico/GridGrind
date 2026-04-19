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

cleanup() {
    rm -f "${sentinel_path}"
}

trap cleanup EXIT

[[ ! -e "${sentinel_path}" ]] || die "sentinel path already exists: ${sentinel_path}"

cat >"${sentinel_path}" <<'EOF'
package dev.erst.gridgrind.cli;

import java.util.*;

final class WildcardImportGateSentinel {}
EOF

set +e
output="$("${repo_root}/gradlew" --console=plain verifyExplicitImports 2>&1)"
status=$?
set -e

[[ ${status} -ne 0 ]] || die 'verifyExplicitImports unexpectedly accepted a wildcard import in src/main'
printf '%s\n' "${output}" | grep -F "${sentinel_relative_path}:3: import java.util.*;" >/dev/null \
    || die 'verifyExplicitImports did not report the injected wildcard import precisely'

printf 'explicit-import gate regression: success\n'
