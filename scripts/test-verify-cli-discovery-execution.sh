#!/usr/bin/env bash
# Exercise the CLI discovery execution verifier against fake jar and binary launchers so long-running
# published-example and recipe verification keeps emitting operator-visible progress.

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
readonly verify_script="${repo_root}/scripts/verify-cli-discovery-execution.sh"

[[ -x "${verify_script}" ]] || die "missing executable verifier script at ${verify_script}"

test_root="$(mktemp -d)"
cleanup() {
    rm -rf "${test_root}"
}
trap cleanup EXIT

readonly fake_bin_dir="${test_root}/bin"
readonly fake_java="${fake_bin_dir}/java"
readonly fake_binary="${fake_bin_dir}/gridgrind"
readonly fake_jar="${test_root}/gridgrind.jar"
mkdir -p "${fake_bin_dir}"
printf 'fake jar\n' > "${fake_jar}"

cat > "${fake_java}" <<'EOF'
#!/usr/bin/env bash
set -euo pipefail

[[ "${1:-}" == "-jar" ]] || {
    printf 'unexpected java invocation: %s\n' "$*" >&2
    exit 1
}
shift 2

command_name="${1:-}"
shift || true

response_path=''
request_path=''

while [[ $# -gt 0 ]]; do
    case "${1}" in
        --response)
            response_path="${2}"
            shift 2
            ;;
        --request)
            request_path="${2}"
            shift 2
            ;;
        --lookup)
            shift 2
            ;;
        *)
            shift
            ;;
    esac
done

write_json_file() {
    local target_path=$1
    local payload=$2
    mkdir -p "$(dirname -- "${target_path}")"
    printf '%s\n' "${payload}" > "${target_path}"
}

case "${command_name}" in
    --print-recipe-catalog)
        cat <<'JSON'
{"protocolVersion":"V2","recipes":[{"view":"EXAMPLE","id":"BUDGET","requestFileName":"budget-request.json","summary":"Budget example.","workspaceMode":"SELF_CONTAINED","requiredWorkspacePaths":[]},{"view":"TASK_STARTER","id":"DASHBOARD","requestFileName":"dashboard-request.json","summary":"Dashboard starter.","workspaceMode":"SELF_CONTAINED","requiredWorkspacePaths":[]}]}
JSON
        ;;
    --print-recipe)
        [[ -n "${response_path}" ]] || {
            printf 'missing response path for --print-recipe\n' >&2
            exit 1
        }
        write_json_file "${response_path}" '{"persistence":{"type":"SAVE_AS","path":"generated-workbooks/fake.xlsx"}}'
        ;;
    --doctor-request)
        [[ -n "${response_path}" ]] || {
            printf 'missing response path for --doctor-request\n' >&2
            exit 1
        }
        [[ -d "$(dirname -- "${request_path}")/generated-workbooks" ]] || {
            printf 'missing prepared SAVE_AS parent directory\n' >&2
            exit 1
        }
        write_json_file "${response_path}" '{"valid":true}'
        ;;
    --request)
        [[ -n "${response_path}" ]] || {
            printf 'missing response path for --request\n' >&2
            exit 1
        }
        sleep 2
        write_json_file "${response_path}" '{"status":"ok"}'
        ;;
    *)
        printf 'unexpected fake gridgrind command: %s\n' "${command_name}" >&2
        exit 1
        ;;
esac
EOF
chmod +x "${fake_java}"

cat > "${fake_binary}" <<'EOF'
#!/usr/bin/env bash
set -euo pipefail

command_name="${1:-}"
shift || true

response_path=''
request_path=''

while [[ $# -gt 0 ]]; do
    case "${1}" in
        --response)
            response_path="${2}"
            shift 2
            ;;
        --request)
            request_path="${2}"
            shift 2
            ;;
        --lookup)
            shift 2
            ;;
        *)
            shift
            ;;
    esac
done

write_json_file() {
    local target_path=$1
    local payload=$2
    mkdir -p "$(dirname -- "${target_path}")"
    printf '%s\n' "${payload}" > "${target_path}"
}

case "${command_name}" in
    --print-recipe-catalog)
        cat <<'JSON'
{"protocolVersion":"V2","recipes":[{"view":"EXAMPLE","id":"BUDGET","requestFileName":"budget-request.json","summary":"Budget example.","workspaceMode":"SELF_CONTAINED","requiredWorkspacePaths":[]},{"view":"TASK_STARTER","id":"DASHBOARD","requestFileName":"dashboard-request.json","summary":"Dashboard starter.","workspaceMode":"SELF_CONTAINED","requiredWorkspacePaths":[]}]}
JSON
        ;;
    --print-recipe)
        [[ -n "${response_path}" ]] || {
            printf 'missing response path for --print-recipe\n' >&2
            exit 1
        }
        write_json_file "${response_path}" '{"persistence":{"type":"SAVE_AS","path":"generated-workbooks/fake.xlsx"}}'
        ;;
    --doctor-request)
        [[ -n "${response_path}" ]] || {
            printf 'missing response path for --doctor-request\n' >&2
            exit 1
        }
        [[ -d "$(dirname -- "${request_path}")/generated-workbooks" ]] || {
            printf 'missing prepared SAVE_AS parent directory\n' >&2
            exit 1
        }
        write_json_file "${response_path}" '{"valid":true}'
        ;;
    --request)
        [[ -n "${response_path}" ]] || {
            printf 'missing response path for --request\n' >&2
            exit 1
        }
        sleep 2
        write_json_file "${response_path}" '{"status":"ok"}'
        ;;
    *)
        printf 'unexpected fake gridgrind command: %s\n' "${command_name}" >&2
        exit 1
        ;;
esac
EOF
chmod +x "${fake_binary}"

assert_progress_output() {
    local output=$1
    local mode_label=$2

    printf '%s\n' "${output}" | grep -Fq 'Discovery execution examples 1/1: BUDGET printing request' || die \
        "${mode_label}: discovery verifier no longer reports example request printing progress"
    printf '%s\n' "${output}" | grep -Fq 'Discovery execution examples 1/1: BUDGET copying required assets' || die \
        "${mode_label}: discovery verifier no longer reports example asset-copy progress"
    printf '%s\n' "${output}" | grep -Fq 'Discovery execution examples 1/1: BUDGET doctoring request' || die \
        "${mode_label}: discovery verifier no longer reports example doctor progress"
    printf '%s\n' "${output}" | grep -Fq 'Discovery execution examples 1/1: BUDGET executing request' || die \
        "${mode_label}: discovery verifier no longer reports example execution progress"
    printf '%s\n' "${output}" | grep -Fq 'Discovery execution examples 1/1: BUDGET executing request (still running after ' || die \
        "${mode_label}: discovery verifier no longer emits example execution heartbeats during slow runs"
    printf '%s\n' "${output}" | grep -Fq 'Discovery execution examples 1/1: BUDGET succeeded' || die \
        "${mode_label}: discovery verifier no longer reports example success progress"
    printf '%s\n' "${output}" | grep -Fq 'Discovery execution task starters 1/1: DASHBOARD printing request' || die \
        "${mode_label}: discovery verifier no longer reports task-starter request printing progress"
    printf '%s\n' "${output}" | grep -Fq 'Discovery execution task starters 1/1: DASHBOARD copying required assets' || die \
        "${mode_label}: discovery verifier no longer reports task-starter asset-copy progress"
    printf '%s\n' "${output}" | grep -Fq 'Discovery execution task starters 1/1: DASHBOARD doctoring request' || die \
        "${mode_label}: discovery verifier no longer reports task-starter doctor progress"
    printf '%s\n' "${output}" | grep -Fq 'Discovery execution task starters 1/1: DASHBOARD executing request' || die \
        "${mode_label}: discovery verifier no longer reports task-starter execution progress"
    printf '%s\n' "${output}" | grep -Fq 'Discovery execution task starters 1/1: DASHBOARD executing request (still running after ' || die \
        "${mode_label}: discovery verifier no longer emits task-starter execution heartbeats during slow runs"
    printf '%s\n' "${output}" | grep -Fq 'Discovery execution task starters 1/1: DASHBOARD succeeded' || die \
        "${mode_label}: discovery verifier no longer reports task-starter success progress"
}

jar_output="$(
    PATH="${fake_bin_dir}:${PATH}" \
        GRIDGRIND_DISCOVERY_EXECUTION_HEARTBEAT_SECONDS=1 \
        "${verify_script}" jar "${fake_jar}"
)"
assert_progress_output "${jar_output}" "jar mode"

binary_output="$(
    GRIDGRIND_DISCOVERY_EXECUTION_HEARTBEAT_SECONDS=1 \
        "${verify_script}" binary "${fake_binary}"
)"
assert_progress_output "${binary_output}" "binary mode"

printf 'verify-cli-discovery-execution regression: success\n'
