#!/usr/bin/env bash
# Guard the documented Jazzer wrapper invocation shape: Gradle global options and properties must
# be forwarded before the task name so commands like
#   jazzer/bin/fuzz-protocol-request -PjazzerMaxDuration=5m --console=plain
# really fuzz instead of falling back to Gradle help.

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
readonly generic_args_log="${tmp_dir}/status-invocation.log"
readonly active_args_log="${tmp_dir}/fuzz-invocation.log"
readonly summary_args_log="${tmp_dir}/summary-invocation.log"

mkdir -p "${fake_jazzer_bin_dir}" "${fake_scripts_dir}" "${fake_repo_root}/jazzer"
cp "${source_run_task}" "${fake_jazzer_bin_dir}/_run-task"
cp "${source_run_lock_support}" "${fake_jazzer_bin_dir}/_run-lock-support"
cp "${source_repo_lock_support}" "${fake_scripts_dir}/repo-verification-lock-support.sh"
chmod +x "${fake_jazzer_bin_dir}/_run-task" "${fake_jazzer_bin_dir}/_run-lock-support"

cat > "${fake_repo_root}/gradlew" <<'EOF'
#!/usr/bin/env bash
set -euo pipefail

repo_root="$(cd -P -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
generic_args_log="${repo_root}/../status-invocation.log"
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

printf '%s\n' "$@" > "${generic_args_log}"

exit 0
EOF
chmod +x "${fake_repo_root}/gradlew"

bash "${fake_jazzer_bin_dir}/_run-task" jazzerStatus -PjazzerTarget=protocol-request --console=plain

mapfile -t generic_args < "${generic_args_log}"
expected_generic=(
    --project-dir
    "${fake_repo_root}/jazzer"
    -PjazzerTarget=protocol-request
    --console=plain
    jazzerStatus
)

[[ ${#generic_args[@]} -eq ${#expected_generic[@]} ]] || die \
    "generic Jazzer wrapper invocation count changed unexpectedly"
for index in "${!expected_generic[@]}"; do
    [[ "${generic_args[$index]}" == "${expected_generic[$index]}" ]] || die \
        "generic Jazzer wrapper invocation reordered Gradle arguments unexpectedly"
done

bash "${fake_jazzer_bin_dir}/_run-task" fuzzProtocolRequest -PjazzerMaxDuration=30s --console=plain

mapfile -t active_args < "${active_args_log}"
expected_active=(
    --project-dir
    "${fake_repo_root}/jazzer"
    --no-daemon
    -PjazzerMaxDuration=30s
    --console=plain
    fuzzProtocolRequest
)

[[ ${#active_args[@]} -eq ${#expected_active[@]} ]] || die \
    "active Jazzer wrapper invocation count changed unexpectedly"
for index in "${!expected_active[@]}"; do
    [[ "${active_args[$index]}" == "${expected_active[$index]}" ]] || die \
        "active Jazzer wrapper invocation reordered Gradle arguments unexpectedly"
done

mapfile -t summary_args < "${summary_args_log}"
[[ " ${summary_args[*]} " == *" jazzerSummarizeRun "* ]] || die \
    "active Jazzer wrapper no longer runs the summary task after fuzz execution"

printf 'jazzer-wrapper-arg-order regression: success\n'
