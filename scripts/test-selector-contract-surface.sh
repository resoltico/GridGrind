#!/usr/bin/env bash
# Keep the public docs/examples surface and release-authored request fixtures on the canonical
# step, selector, and source-backed contract.

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

check_no_matches() {
    local pattern=$1
    local description=$2
    shift 2
    if rg -n "${pattern}" "$@" >/dev/null; then
        die "${description}"
    fi
}

check_no_pcre_matches() {
    local pattern=$1
    local description=$2
    shift 2
    if rg -nUP "${pattern}" "$@" >/dev/null; then
        die "${description}"
    fi
}

readonly public_surface=(
    "${repo_root}/README.md"
    "${repo_root}/docs"
    "${repo_root}/examples"
)
readonly release_request_surface=(
    "${repo_root}/scripts/docker-smoke.sh"
)

check_no_matches '"selection"' \
    'public docs/examples still mention the deleted request field "selection"' \
    "${public_surface[@]}"
check_no_matches '"sourceSheetName"' \
    'public docs/examples still mention the deleted request field "sourceSheetName"' \
    "${public_surface[@]}"
check_no_matches '"type"[[:space:]]*:[[:space:]]*"SELECTED"' \
    'public docs/examples still mention the deleted selector discriminator "SELECTED"' \
    "${public_surface[@]}"
check_no_matches '"type"[[:space:]]*:[[:space:]]*"ALL_USED_CELLS"' \
    'public docs/examples still mention the deleted selector discriminator "ALL_USED_CELLS"' \
    "${public_surface[@]}"
check_no_matches '\| `selection` \|' \
    'public docs still document the deleted request field `selection`' \
    "${repo_root}/README.md" "${repo_root}/docs"
check_no_matches '\| `sourceSheetName` \|' \
    'public docs still document the deleted request field `sourceSheetName`' \
    "${repo_root}/README.md" "${repo_root}/docs"
check_no_matches 'Selection snippets:' \
    'public docs still label selector examples as selection snippets' \
    "${repo_root}/README.md" "${repo_root}/docs"
check_no_matches 'Cell-selection payloads use:|Range-selection payloads use:|Table-selection payloads use:|Sheet-selection payloads use:' \
    'public docs still describe selector payload families with the deleted selection terminology' \
    "${repo_root}/README.md" "${repo_root}/docs"
check_no_matches '"operations"[[:space:]]*:' \
    'public docs/examples still mention the deleted request field "operations"' \
    "${public_surface[@]}"
check_no_matches '"reads"[[:space:]]*:' \
    'public docs/examples still mention the deleted request field "reads"' \
    "${public_surface[@]}"
check_no_matches '"requestId"[[:space:]]*:' \
    'public docs/examples still mention the deleted inspection field "requestId"' \
    "${public_surface[@]}"
check_no_matches '"selector"[[:space:]]*:' \
    'public docs/examples still mention the deleted step field "selector"' \
    "${public_surface[@]}"
check_no_pcre_matches '(?s)"type"\s*:\s*"ENSURE_SHEET"\s*,\s*"sheetName"' \
    'release-surface shell requests still author ENSURE_SHEET with the deleted sheetName field' \
    "${release_request_surface[@]}"
check_no_pcre_matches '(?s)"type"\s*:\s*"APPEND_ROW"\s*,\s*"sheetName"' \
    'release-surface shell requests still author APPEND_ROW with the deleted sheetName field' \
    "${release_request_surface[@]}"
check_no_pcre_matches '(?s)"type"\s*:\s*"TEXT"\s*,\s*"text"' \
    'release-surface shell requests still author text cell values with the deleted inline text field instead of source-backed payloads' \
    "${release_request_surface[@]}"
check_no_pcre_matches '(?s)"type"\s*:\s*"FORMULA"\s*,\s*"formula"' \
    'release-surface shell requests still author formula cell values with the deleted inline formula field instead of source-backed payloads' \
    "${release_request_surface[@]}"
check_no_matches '"operations"[[:space:]]*:' \
    'release-surface shell requests still author the deleted operations array' \
    "${release_request_surface[@]}"
check_no_matches '"reads"[[:space:]]*:' \
    'release-surface shell requests still author the deleted reads array' \
    "${release_request_surface[@]}"
check_no_matches '"requestId"[[:space:]]*:' \
    'release-surface shell requests still author the deleted requestId field' \
    "${release_request_surface[@]}"
check_no_matches '"selector"[[:space:]]*:' \
    'release-surface shell requests still author the deleted selector field' \
    "${release_request_surface[@]}"

printf 'selector-contract-surface regression: success\n'
