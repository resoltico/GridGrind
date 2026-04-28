#!/usr/bin/env bash
# Reproduce and guard portable file-size detection across BSD stat, GNU stat, and wc fallback.

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

assert_file_size_mode() {
    local stat_mode=$1
    local expected_size_value=$2
    local fake_bin_dir_path=$3
    local support_script_path=$4
    local target_file_path=$5
    local fake_stat="${fake_bin_dir_path}/stat"

    case "${stat_mode}" in
        bsd)
            cat > "${fake_stat}" <<'EOF'
#!/usr/bin/env bash
set -euo pipefail
[[ "${1:-}" == "-f" && "${2:-}" == "%z" ]] || exit 1
wc -c < "${3:?}" | tr -d '[:space:]'
EOF
            ;;
        gnu)
            cat > "${fake_stat}" <<'EOF'
#!/usr/bin/env bash
set -euo pipefail
[[ "${1:-}" == "-c" && "${2:-}" == "%s" ]] || exit 1
wc -c < "${3:?}" | tr -d '[:space:]'
EOF
            ;;
        fail)
            cat > "${fake_stat}" <<'EOF'
#!/usr/bin/env bash
exit 1
EOF
            ;;
        *)
            die "unsupported fake stat mode ${stat_mode}"
            ;;
    esac
    chmod +x "${fake_stat}"

    local observed_size=''
    observed_size="$(
        PATH="${fake_bin_dir_path}:$PATH" bash -lc '
            set -euo pipefail
            source "$1"
            file_size_bytes "$2"
        ' bash "${support_script_path}" "${target_file_path}"
    )"
    [[ "${observed_size}" == "${expected_size_value}" ]] || die \
        "file_size_bytes returned ${observed_size} for mode ${stat_mode}, expected ${expected_size_value}"
}

readonly script_dir="$(resolve_script_dir)"
readonly repo_root="$(cd -P -- "${script_dir}/.." && pwd)"
readonly file_support="${repo_root}/scripts/check-file-support.sh"

[[ -f "${file_support}" ]] || die "missing file support helper"

# shellcheck source=/dev/null
source "${file_support}"

tmp_dir="$(mktemp -d)"
cleanup() {
    rm -rf "${tmp_dir}"
}
trap cleanup EXIT

readonly target_file="${tmp_dir}/probe.log"
printf 'abcdef' > "${target_file}"
readonly expected_size='6'
readonly fake_bin_dir="${tmp_dir}/bin"
mkdir -p "${fake_bin_dir}"

assert_file_size_mode bsd "${expected_size}" "${fake_bin_dir}" "${file_support}" "${target_file}"
assert_file_size_mode gnu "${expected_size}" "${fake_bin_dir}" "${file_support}" "${target_file}"
assert_file_size_mode fail "${expected_size}" "${fake_bin_dir}" "${file_support}" "${target_file}"

[[ "$(file_size_bytes "${tmp_dir}/missing.log")" == '0' ]] || die "missing files must report size 0"

printf 'check-file-support regression: success\n'
