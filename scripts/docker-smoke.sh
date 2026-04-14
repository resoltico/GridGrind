#!/usr/bin/env bash
# Build the local Docker image and verify the packaged CLI works correctly from a non-default
# working directory with weird request, response, and workbook paths, including reopening an
# existing workbook and low-memory streaming write readback from the materialized output.

set -euo pipefail

die() {
    printf 'error: %s\n' "$1" >&2
    exit 1
}

require_match() {
    local text=$1
    local pattern=$2
    local message=$3

    if ! printf '%s\n' "${text}" | grep -Eq "${pattern}"; then
        die "${message}"
    fi
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
readonly image_tag="gridgrind-docker-smoke:$$"
readonly smoke_root="${repo_root}/tmp/docker smoke.$$"
readonly docker_run_user="$(id -u):$(id -g)"
anonymous_docker_config=''
docker_endpoint=''

resolve_docker_buildx_plugin() {
    local docker_binary=''
    local -a candidates=()
    local candidate=''

    if candidate="$(command -v docker-buildx 2>/dev/null || true)"; then
        if [[ -n "${candidate}" ]]; then
            candidates+=("${candidate}")
        fi
    fi

    docker_binary="$(command -v docker)"
    candidates+=(
        "${HOME}/.docker/cli-plugins/docker-buildx"
        "/Applications/Docker.app/Contents/Resources/cli-plugins/docker-buildx"
        "/usr/local/lib/docker/cli-plugins/docker-buildx"
        "/usr/local/libexec/docker/cli-plugins/docker-buildx"
        "/opt/homebrew/lib/docker/cli-plugins/docker-buildx"
        "/opt/homebrew/libexec/docker/cli-plugins/docker-buildx"
        "/usr/lib/docker/cli-plugins/docker-buildx"
        "/usr/libexec/docker/cli-plugins/docker-buildx"
        "/usr/lib64/docker/cli-plugins/docker-buildx"
        "/usr/share/docker/cli-plugins/docker-buildx"
        "$(cd -P -- "$(dirname -- "${docker_binary}")" && pwd)/docker-buildx"
    )

    for candidate in "${candidates[@]}"; do
        if [[ -n "${candidate}" && -x "${candidate}" ]]; then
            printf '%s\n' "${candidate}"
            return 0
        fi
    done

    return 1
}

docker_with_repo_config() {
    if [[ -n "${docker_endpoint}" ]]; then
        DOCKER_CONFIG="${anonymous_docker_config}" DOCKER_HOST="${docker_endpoint}" docker "$@"
        return
    fi
    DOCKER_CONFIG="${anonymous_docker_config}" docker "$@"
}

cleanup() {
    local exit_code=$?
    # Mounted-path artifacts should stay caller-owned, but keep sudo as a defensive cleanup fallback.
    rm -rf "${smoke_root}" || sudo rm -rf "${smoke_root}" || true
    if command -v docker >/dev/null 2>&1 && [[ -n "${anonymous_docker_config}" ]]; then
        docker_with_repo_config image rm -f "${image_tag}" >/dev/null 2>&1 || true
        rm -rf "${anonymous_docker_config}" || true
    fi
    exit "${exit_code}"
}

trap cleanup EXIT

command -v docker >/dev/null 2>&1 || die "docker is required for the Docker smoke gate"
docker buildx version >/dev/null 2>&1 || die "docker buildx is required for the Docker smoke gate"
[[ -f "${repo_root}/Dockerfile" ]] || die "missing Dockerfile at ${repo_root}/Dockerfile"
[[ -f "${repo_root}/cli/build/libs/gridgrind.jar" ]] || die \
    "missing CLI fat JAR at ${repo_root}/cli/build/libs/gridgrind.jar; run ./gradlew :cli:shadowJar first"

docker_endpoint="${DOCKER_HOST:-}"
if [[ -z "${docker_endpoint}" ]]; then
    docker_endpoint="$(
        docker context inspect "$(docker context show 2>/dev/null || true)" \
            --format '{{.Endpoints.docker.Host}}' 2>/dev/null || true
    )"
fi
anonymous_docker_config="$(mktemp -d "${TMPDIR:-/tmp}/gridgrind-docker-config.XXXXXX")"
printf '{}\n' > "${anonymous_docker_config}/config.json"

if ! docker_with_repo_config buildx version >/dev/null 2>&1; then
    docker_buildx_plugin="$(resolve_docker_buildx_plugin)" || die \
        "docker buildx is available in the current shell, but no reusable docker-buildx plugin binary was found for the anonymous DOCKER_CONFIG"
    mkdir -p "${anonymous_docker_config}/cli-plugins"
    ln -s "${docker_buildx_plugin}" "${anonymous_docker_config}/cli-plugins/docker-buildx"
    docker_with_repo_config buildx version >/dev/null 2>&1 || die \
        "docker buildx is not reachable through the anonymous DOCKER_CONFIG even after staging ${docker_buildx_plugin}"
fi

mkdir -p "${smoke_root}/requests odd"

readonly request_rel='requests odd/request [docker #smoke].json'
readonly response_rel='responses odd/nested/response [docker #smoke].json'
readonly workbook_rel='books odd/nested/office [docker #smoke].xlsx'
readonly existing_request_rel='requests odd/request reopen [docker #smoke].json'
readonly existing_response_rel='responses odd/nested/reopen [docker #smoke].json'
readonly streaming_request_rel='requests odd/request streaming [docker #smoke].json'
readonly streaming_response_rel='responses odd/nested/streaming [docker #smoke].json'
readonly streaming_workbook_rel='books odd/nested/office streaming [docker #smoke].xlsx'
readonly request_path="${smoke_root}/${request_rel}"
readonly response_path="${smoke_root}/${response_rel}"
readonly workbook_path="${smoke_root}/${workbook_rel}"
readonly existing_request_path="${smoke_root}/${existing_request_rel}"
readonly existing_response_path="${smoke_root}/${existing_response_rel}"
readonly streaming_request_path="${smoke_root}/${streaming_request_rel}"
readonly streaming_response_path="${smoke_root}/${streaming_response_rel}"
readonly streaming_workbook_path="${smoke_root}/${streaming_workbook_rel}"
readonly create_stderr_path="${smoke_root}/stderr create [docker #smoke].log"
readonly existing_stderr_path="${smoke_root}/stderr reopen [docker #smoke].log"
readonly streaming_stderr_path="${smoke_root}/stderr streaming [docker #smoke].log"

cat > "${request_path}" <<JSON
{
  "source": { "type": "NEW" },
  "persistence": {
    "type": "SAVE_AS",
    "path": "${workbook_rel}"
  },
  "operations": [
    { "type": "ENSURE_SHEET", "sheetName": "Smoke" }
  ],
  "reads": [
    { "type": "GET_WORKBOOK_SUMMARY", "requestId": "workbook" }
  ]
}
JSON

cat > "${existing_request_path}" <<JSON
{
  "source": {
    "type": "EXISTING",
    "path": "${workbook_rel}"
  },
  "persistence": { "type": "NONE" },
  "reads": [
    { "type": "GET_WORKBOOK_SUMMARY", "requestId": "existing-workbook" }
  ]
}
JSON

cat > "${streaming_request_path}" <<JSON
{
  "source": { "type": "NEW" },
  "persistence": {
    "type": "SAVE_AS",
    "path": "${streaming_workbook_rel}"
  },
  "executionMode": {
    "readMode": "EVENT_READ",
    "writeMode": "STREAMING_WRITE"
  },
  "operations": [
    { "type": "ENSURE_SHEET", "sheetName": "Ledger" },
    {
      "type": "APPEND_ROW",
      "sheetName": "Ledger",
      "values": [
        { "type": "TEXT", "text": "Team" },
        { "type": "TEXT", "text": "Task" },
        { "type": "TEXT", "text": "Hours" }
      ]
    },
    {
      "type": "APPEND_ROW",
      "sheetName": "Ledger",
      "values": [
        { "type": "TEXT", "text": "Ops" },
        { "type": "TEXT", "text": "Badge prep" },
        { "type": "NUMBER", "number": 6.5 }
      ]
    }
  ],
  "reads": [
    { "type": "GET_WORKBOOK_SUMMARY", "requestId": "streaming-workbook" },
    { "type": "GET_SHEET_SUMMARY", "requestId": "streaming-sheet", "sheetName": "Ledger" }
  ]
}
JSON

printf 'Docker smoke: building local image\n'
docker_with_repo_config buildx build --load -t "${image_tag}" "${repo_root}" >/dev/null

printf 'Docker smoke: verifying custom workdir and weird paths\n'
help_output="$(docker_with_repo_config run --rm \
    --user "${docker_run_user}" \
    -w /workdir \
    -v "${smoke_root}:/workdir" \
    "${image_tag}" \
    --help | tr -d '\r')"
require_match "${help_output}" '^GridGrind ' \
    "docker smoke help output did not include the banner"
require_match "${help_output}" '^Usage:' \
    "docker smoke help output did not include the usage section"

version_output="$(docker_with_repo_config run --rm \
    --user "${docker_run_user}" \
    -w /workdir \
    -v "${smoke_root}:/workdir" \
    "${image_tag}" \
    --version | tr -d '\r')"
require_match "${version_output}" '^GridGrind ' \
    "docker smoke version output did not include the application name"

docker_with_repo_config run --rm \
    --user "${docker_run_user}" \
    -w /workdir \
    -v "${smoke_root}:/workdir" \
    "${image_tag}" \
    --request "${request_rel}" \
    --response "${response_rel}" >/dev/null 2>"${create_stderr_path}"

[[ -f "${response_path}" ]] || die "docker smoke response file was not written: ${response_path}"
[[ -f "${workbook_path}" ]] || die "docker smoke workbook file was not written: ${workbook_path}"
grep -Eq '"status"[[:space:]]*:[[:space:]]*"SUCCESS"' "${response_path}" || die \
    "docker smoke response did not report SUCCESS"
[[ ! -s "${create_stderr_path}" ]] || die \
    "docker smoke create request wrote unexpected stderr: $(tr '\n' ' ' < "${create_stderr_path}")"

printf 'Docker smoke: reopening saved workbook through EXISTING source\n'
docker_with_repo_config run --rm \
    --user "${docker_run_user}" \
    -w /workdir \
    -v "${smoke_root}:/workdir" \
    "${image_tag}" \
    --request "${existing_request_rel}" \
    --response "${existing_response_rel}" >/dev/null 2>"${existing_stderr_path}"

[[ -f "${existing_response_path}" ]] || die \
    "docker smoke reopen response file was not written: ${existing_response_path}"
grep -Eq '"status"[[:space:]]*:[[:space:]]*"SUCCESS"' "${existing_response_path}" || die \
    "docker smoke EXISTING-source reopen did not report SUCCESS"
[[ ! -s "${existing_stderr_path}" ]] || die \
    "docker smoke EXISTING-source reopen wrote unexpected stderr: $(tr '\n' ' ' < "${existing_stderr_path}")"

printf 'Docker smoke: verifying STREAMING_WRITE readback from materialized output\n'
docker_with_repo_config run --rm \
    --user "${docker_run_user}" \
    -w /workdir \
    -v "${smoke_root}:/workdir" \
    "${image_tag}" \
    --request "${streaming_request_rel}" \
    --response "${streaming_response_rel}" >/dev/null 2>"${streaming_stderr_path}"

[[ -f "${streaming_response_path}" ]] || die \
    "docker smoke streaming response file was not written: ${streaming_response_path}"
[[ -f "${streaming_workbook_path}" ]] || die \
    "docker smoke streaming workbook file was not written: ${streaming_workbook_path}"
grep -Eq '"status"[[:space:]]*:[[:space:]]*"SUCCESS"' "${streaming_response_path}" || die \
    "docker smoke STREAMING_WRITE readback did not report SUCCESS"
[[ ! -s "${streaming_stderr_path}" ]] || die \
    "docker smoke STREAMING_WRITE readback wrote unexpected stderr: $(tr '\n' ' ' < "${streaming_stderr_path}")"

printf 'Docker smoke: success\n'
