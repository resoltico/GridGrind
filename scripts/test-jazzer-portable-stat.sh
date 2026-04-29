#!/usr/bin/env bash
# Guard the active Jazzer wrapper against host/guest stat syntax drift. The terminal-only Docker
# path runs under GNU coreutils, while host-native macOS uses BSD stat.

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
readonly source_run_task="${repo_root}/jazzer/bin/_run-task"
readonly source_run_lock_support="${repo_root}/jazzer/bin/_run-lock-support"
readonly source_repo_lock_support="${repo_root}/scripts/repo-verification-lock-support.sh"

[[ -f "${source_run_task}" ]] || die "missing ${source_run_task}"
[[ -f "${source_run_lock_support}" ]] || die "missing ${source_run_lock_support}"
[[ -f "${source_repo_lock_support}" ]] || die "missing ${source_repo_lock_support}"

tmp_dir="$(mktemp -d)"
tmp_dir="$(cd -P -- "${tmp_dir}" && pwd)"
cleanup() {
    rm -rf "${tmp_dir}"
}
trap cleanup EXIT

readonly fake_repo_root="${tmp_dir}/repo"
readonly fake_jazzer_bin_dir="${fake_repo_root}/jazzer/bin"
readonly fake_scripts_dir="${fake_repo_root}/scripts"
readonly fake_tools_dir="${tmp_dir}/fake-tools"
readonly active_args_log="${tmp_dir}/fuzz-invocation.log"
readonly summary_args_log="${tmp_dir}/summary-invocation.log"

mkdir -p \
    "${fake_jazzer_bin_dir}" \
    "${fake_scripts_dir}" \
    "${fake_tools_dir}" \
    "${fake_repo_root}/jazzer/.local/runs/protocol-request/.cifuzz-corpus/seed"

printf 'seed-bytes\n' > "${fake_repo_root}/jazzer/.local/runs/protocol-request/.cifuzz-corpus/seed/input.bin"

cp "${source_run_task}" "${fake_jazzer_bin_dir}/_run-task"
cp "${source_run_lock_support}" "${fake_jazzer_bin_dir}/_run-lock-support"
cp "${source_repo_lock_support}" "${fake_scripts_dir}/repo-verification-lock-support.sh"
chmod +x "${fake_jazzer_bin_dir}/_run-task" "${fake_jazzer_bin_dir}/_run-lock-support"

cat > "${fake_tools_dir}/stat" <<'EOF'
#!/usr/bin/env bash
set -euo pipefail

if [[ $# -ne 3 ]]; then
    exit 1
fi

mode="$1"
format="$2"
path="$3"

if [[ "${mode}" == "-f" ]]; then
    exit 1
fi

if [[ "${mode}" != "-c" || "${format}" != "%s" ]]; then
    exit 1
fi

wc -c < "${path}" | tr -d '[:space:]'
printf '\n'
EOF
chmod +x "${fake_tools_dir}/stat"

cat > "${fake_repo_root}/gradlew" <<'EOF'
#!/usr/bin/env bash
set -euo pipefail

repo_root="$(cd -P -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
active_args_log="${repo_root}/../fuzz-invocation.log"
summary_args_log="${repo_root}/../summary-invocation.log"

for arg in "$@"; do
    if [[ "${arg}" == "jazzerSummarizeRun" ]]; then
        printf '%s\n' "$@" > "${summary_args_log}"
        exit 0
    fi
done

for arg in "$@"; do
    if [[ "${arg}" == "fuzzProtocolRequest" ]]; then
        printf '%s\n' "$@" > "${active_args_log}"
        exit 0
    fi
done

exit 0
EOF
chmod +x "${fake_repo_root}/gradlew"

PATH="${fake_tools_dir}:$PATH" bash "${fake_jazzer_bin_dir}/_run-task" fuzzProtocolRequest -PjazzerMaxDuration=30s --console=plain

[[ -f "${active_args_log}" ]] || die \
    "active Jazzer wrapper failed before invoking Gradle under GNU-style stat semantics"
[[ -f "${summary_args_log}" ]] || die \
    "active Jazzer wrapper failed before summarizing under GNU-style stat semantics"

printf 'jazzer-portable-stat regression: success\n'
