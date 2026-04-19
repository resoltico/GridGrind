#!/usr/bin/env bash
# Rewrite the committed JSON example fixtures from the contract-owned built-in example registry.

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
readonly examples_dir="${repo_root}/examples"

command -v python3 >/dev/null 2>&1 || die "python3 is required"
[[ -d "${examples_dir}" ]] || die "missing examples directory at ${examples_dir}"

catalog_json="$(mktemp "${TMPDIR:-/tmp}/gridgrind-example-catalog.XXXXXX.json")"
trap 'rm -f "${catalog_json}"' EXIT

./gradlew -q :cli:run --args="--print-protocol-catalog" > "${catalog_json}"

mapfile -t example_pairs < <(
    python3 - "${catalog_json}" <<'PY'
import json
import sys
from pathlib import Path

catalog = json.loads(Path(sys.argv[1]).read_text())
for example in catalog["shippedExamples"]:
    print(example["id"] + "\t" + example["fileName"])
PY
)

[[ ${#example_pairs[@]} -gt 0 ]] || die "protocol catalog returned no shippedExamples"

declare -A generated_files=()
for pair in "${example_pairs[@]}"; do
    example_id="${pair%%$'\t'*}"
    file_name="${pair#*$'\t'}"
    target_path="${examples_dir}/${file_name}"
    ./gradlew -q :cli:run --args="--print-example ${example_id}" > "${target_path}"
    generated_files["${file_name}"]=1
done

while IFS= read -r current_file; do
    file_name="$(basename "${current_file}")"
    [[ -n "${generated_files[${file_name}]:-}" ]] && continue
    rm -f "${current_file}"
done < <(find "${examples_dir}" -maxdepth 1 -type f -name '*.json' | sort)

printf 'generated example fixtures refreshed: %s\n' "${#example_pairs[@]}"
