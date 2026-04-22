#!/usr/bin/env bash
# Keep the documented Jazzer wrapper surface real from a clean checkout.

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
readonly jazzer_bin_dir="${repo_root}/jazzer/bin"
readonly jazzer_readme="${repo_root}/jazzer/README.md"
readonly jazzer_operations_doc="${repo_root}/docs/DEVELOPER_JAZZER_OPERATIONS.md"

expected_support_scripts=(
    _run-lock-support
    _run-task
)
expected_public_scripts=(
    clean-local-corpus
    clean-local-findings
    fuzz-all
    fuzz-engine-command-sequence
    fuzz-protocol-request
    fuzz-protocol-workflow
    fuzz-xlsx-roundtrip
    list-corpus
    list-findings
    promote
    refresh-promoted-metadata
    regression
    replay
    report
    status
)
expected_scripts=("${expected_support_scripts[@]}" "${expected_public_scripts[@]}")

[[ -d "${jazzer_bin_dir}" ]] || die "jazzer/bin is missing from a clean checkout"

for script_name in "${expected_scripts[@]}"; do
    script_path="${jazzer_bin_dir}/${script_name}"
    [[ -f "${script_path}" ]] || die "missing tracked Jazzer script ${script_name}"
    [[ -x "${script_path}" ]] || die "Jazzer script ${script_name} is not executable"
    git -C "${repo_root}" ls-files --error-unmatch "jazzer/bin/${script_name}" >/dev/null 2>&1 || die \
        "Jazzer script ${script_name} is not tracked by git"
done

while IFS= read -r script_path; do
    [[ -n "${script_path}" ]] || continue
    script_name="$(basename -- "${script_path}")"
    found=0
    for expected_name in "${expected_scripts[@]}"; do
        if [[ "${expected_name}" == "${script_name}" ]]; then
            found=1
            break
        fi
    done
    [[ ${found} -eq 1 ]] || die "unexpected Jazzer bin entry ${script_name}; update the public-surface inventory"
done < <(find "${jazzer_bin_dir}" -maxdepth 1 -type f | sort)

for script_name in "${expected_public_scripts[@]}"; do
    grep -Fq "jazzer/bin/${script_name}" "${jazzer_readme}" || die \
        "jazzer README no longer documents ${script_name}"
    grep -Fq "jazzer/bin/${script_name}" "${jazzer_operations_doc}" || die \
        "DEVELOPER_JAZZER_OPERATIONS no longer documents ${script_name}"
done

printf 'jazzer-public-surface regression: success\n'
