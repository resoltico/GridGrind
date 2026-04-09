#!/usr/bin/env bash
# Build the local Docker image and verify it runs correctly from a non-default working directory
# with weird request, response, and workbook paths.

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

cleanup() {
    local exit_code=$?
    # Files created by the container (root-owned on Linux) may require sudo to delete in CI.
    rm -rf "${smoke_root}" || sudo rm -rf "${smoke_root}" || true
    docker image rm -f "${image_tag}" >/dev/null 2>&1 || true
    exit "${exit_code}"
}

trap cleanup EXIT

command -v docker >/dev/null 2>&1 || die "docker is required for the Docker smoke gate"
[[ -f "${repo_root}/Dockerfile" ]] || die "missing Dockerfile at ${repo_root}/Dockerfile"
[[ -f "${repo_root}/cli/build/libs/gridgrind.jar" ]] || die \
    "missing CLI fat JAR at ${repo_root}/cli/build/libs/gridgrind.jar; run ./gradlew :cli:shadowJar first"

mkdir -p "${smoke_root}/requests odd"

readonly request_rel='requests odd/request [docker #smoke].json'
readonly response_rel='responses odd/nested/response [docker #smoke].json'
readonly workbook_rel='books odd/nested/office [docker #smoke].xlsx'
readonly request_path="${smoke_root}/${request_rel}"
readonly response_path="${smoke_root}/${response_rel}"
readonly workbook_path="${smoke_root}/${workbook_rel}"

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

printf 'Docker smoke: building local image\n'
docker build -t "${image_tag}" "${repo_root}" >/dev/null

printf 'Docker smoke: verifying custom workdir and weird paths\n'
help_output="$(docker run --rm \
    -w /workdir \
    -v "${smoke_root}:/workdir" \
    "${image_tag}" \
    --help | tr -d '\r')"
require_match "${help_output}" '^GridGrind CLI ' \
    "docker smoke help output did not include the CLI banner"
require_match "${help_output}" '^Usage:' \
    "docker smoke help output did not include the usage section"

version_output="$(docker run --rm \
    -w /workdir \
    -v "${smoke_root}:/workdir" \
    "${image_tag}" \
    --version | tr -d '\r')"
require_match "${version_output}" '^gridgrind ' \
    "docker smoke version output did not include the application name"

docker run --rm \
    -w /workdir \
    -v "${smoke_root}:/workdir" \
    "${image_tag}" \
    --request "${request_rel}" \
    --response "${response_rel}" >/dev/null

[[ -f "${response_path}" ]] || die "docker smoke response file was not written: ${response_path}"
[[ -f "${workbook_path}" ]] || die "docker smoke workbook file was not written: ${workbook_path}"
grep -Eq '"status"[[:space:]]*:[[:space:]]*"SUCCESS"' "${response_path}" || die \
    "docker smoke response did not report SUCCESS"

printf 'Docker smoke: success\n'
