#!/usr/bin/env bash
# Build the local Docker image and verify the packaged CLI works correctly from a non-default
# working directory with weird request, response, and workbook paths, including reopening an
# existing workbook, authoring a signature line, and low-memory streaming write readback from the
# materialized output.

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
readonly smoke_root="$(mktemp -d "${TMPDIR:-/tmp}/gridgrind docker smoke.XXXXXX")"
readonly docker_run_user="$(id -u):$(id -g)"
readonly cli_jar_path="${repo_root}/cli/build/libs/gridgrind.jar"
readonly repo_lock_support="${repo_root}/scripts/repo-verification-lock-support.sh"
readonly lock_dir="${repo_root}/tmp/repo-verification-lock"
readonly pid_file="${lock_dir}/pid"
readonly gradle_user_home="${GRIDGRIND_GRADLE_USER_HOME:-${repo_root}/tmp/gradle-user-home}"
anonymous_docker_config=''
docker_endpoint=''
project_cache_dir=''

prepare_project_cache_dir() {
    local cache_root=$1
    if [[ -z "${cache_root}" ]]; then
        return 0
    fi
    mkdir -p "${cache_root}"
    mktemp -d "${cache_root%/}/docker-smoke.XXXXXX"
}

project_cache_dir="$(prepare_project_cache_dir "${GRIDGRIND_PROJECT_CACHE_DIR:-}")"
[[ -f "${repo_lock_support}" ]] || die "missing repo verification lock helper at ${repo_lock_support}"

# shellcheck source=/dev/null
source "${repo_lock_support}"

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
    if [[ -n "${project_cache_dir}" ]]; then
        rm -rf "${project_cache_dir}" || true
    fi
    if command -v docker >/dev/null 2>&1 && [[ -n "${anonymous_docker_config}" ]]; then
        docker_with_repo_config image rm -f "${image_tag}" >/dev/null 2>&1 || true
        rm -rf "${anonymous_docker_config}" || true
    fi
    cleanup_lock
    exit "${exit_code}"
}

trap cleanup EXIT

command -v docker >/dev/null 2>&1 || die "docker is required for the Docker smoke gate"
docker buildx version >/dev/null 2>&1 || die "docker buildx is required for the Docker smoke gate"
[[ -f "${repo_root}/Dockerfile" ]] || die "missing Dockerfile at ${repo_root}/Dockerfile"
[[ -x "${repo_root}/gradlew" ]] || die "missing Gradle wrapper at ${repo_root}/gradlew"

mkdir -p "${gradle_user_home}"
acquire_lock

printf 'Docker smoke: rebuilding CLI fat JAR\n'
gradle_command=(
    env
    "GRADLE_USER_HOME=${gradle_user_home}"
    "${repo_root}/gradlew"
    --console=plain
    --no-daemon
)
if [[ -n "${project_cache_dir}" ]]; then
    gradle_command+=(--project-cache-dir "${project_cache_dir}")
fi
gradle_command+=(:cli:shadowJar)
"${gradle_command[@]}" >/dev/null
[[ -f "${cli_jar_path}" ]] || die "missing CLI fat JAR at ${cli_jar_path} after :cli:shadowJar"

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
readonly signature_request_rel='requests odd/request signature [docker #smoke].json'
readonly signature_response_rel='responses odd/nested/signature [docker #smoke].json'
readonly signature_workbook_rel='books odd/nested/office signature [docker #smoke].xlsx'
readonly streaming_request_rel='requests odd/request streaming [docker #smoke].json'
readonly streaming_response_rel='responses odd/nested/streaming [docker #smoke].json'
readonly streaming_read_request_rel='requests odd/request streaming readback [docker #smoke].json'
readonly streaming_read_response_rel='responses odd/nested/streaming readback [docker #smoke].json'
readonly streaming_workbook_rel='books odd/nested/office streaming [docker #smoke].xlsx'
readonly request_path="${smoke_root}/${request_rel}"
readonly response_path="${smoke_root}/${response_rel}"
readonly request_dir="${smoke_root}/requests odd"
readonly legacy_workbook_path="${smoke_root}/${workbook_rel}"
readonly workbook_path="${request_dir}/${workbook_rel}"
readonly existing_request_path="${smoke_root}/${existing_request_rel}"
readonly existing_response_path="${smoke_root}/${existing_response_rel}"
readonly signature_request_path="${smoke_root}/${signature_request_rel}"
readonly signature_response_path="${smoke_root}/${signature_response_rel}"
readonly legacy_signature_workbook_path="${smoke_root}/${signature_workbook_rel}"
readonly signature_workbook_path="${request_dir}/${signature_workbook_rel}"
readonly streaming_request_path="${smoke_root}/${streaming_request_rel}"
readonly streaming_response_path="${smoke_root}/${streaming_response_rel}"
readonly streaming_read_request_path="${smoke_root}/${streaming_read_request_rel}"
readonly streaming_read_response_path="${smoke_root}/${streaming_read_response_rel}"
readonly legacy_streaming_workbook_path="${smoke_root}/${streaming_workbook_rel}"
readonly streaming_workbook_path="${request_dir}/${streaming_workbook_rel}"
readonly create_stderr_path="${smoke_root}/stderr create [docker #smoke].log"
readonly existing_stderr_path="${smoke_root}/stderr reopen [docker #smoke].log"
readonly signature_stderr_path="${smoke_root}/stderr signature [docker #smoke].log"
readonly streaming_stderr_path="${smoke_root}/stderr streaming [docker #smoke].log"
readonly streaming_read_stderr_path="${smoke_root}/stderr streaming readback [docker #smoke].log"

cat > "${request_path}" <<JSON
{
  "source": {
    "type": "NEW"
  },
  "persistence": {
    "type": "SAVE_AS",
    "path": "${workbook_rel}"
  },
  "execution": {
    "journal": {
      "level": "VERBOSE"
    }
  },
  "steps": [
    {
      "stepId": "ensure-smoke",
      "target": {
        "type": "SHEET_BY_NAME",
        "name": "Smoke"
      },
      "action": {
        "type": "ENSURE_SHEET"
      }
    },
    {
      "stepId": "workbook",
      "target": {
        "type": "WORKBOOK_CURRENT"
      },
      "query": {
        "type": "GET_WORKBOOK_SUMMARY"
      }
    }
  ]
}
JSON

cat > "${existing_request_path}" <<JSON
{
  "source": {
    "type": "EXISTING",
    "path": "${workbook_rel}"
  },
  "persistence": {
    "type": "NONE"
  },
  "steps": [
    {
      "stepId": "existing-workbook",
      "target": {
        "type": "WORKBOOK_CURRENT"
      },
      "query": {
        "type": "GET_WORKBOOK_SUMMARY"
      }
    }
  ]
}
JSON

cat > "${signature_request_path}" <<JSON
{
  "source": {
    "type": "NEW"
  },
  "persistence": {
    "type": "SAVE_AS",
    "path": "${signature_workbook_rel}"
  },
  "steps": [
    {
      "stepId": "ensure-approvals",
      "target": {
        "type": "SHEET_BY_NAME",
        "name": "Approvals"
      },
      "action": {
        "type": "ENSURE_SHEET"
      }
    },
    {
      "stepId": "set-signature-line",
      "target": {
        "type": "SHEET_BY_NAME",
        "name": "Approvals"
      },
      "action": {
        "type": "SET_SIGNATURE_LINE",
        "signatureLine": {
          "name": "BudgetSignature",
          "anchor": {
            "type": "TWO_CELL",
            "from": {
              "columnIndex": 1,
              "rowIndex": 1,
              "dx": 0,
              "dy": 0
            },
            "to": {
              "columnIndex": 4,
              "rowIndex": 6,
              "dx": 0,
              "dy": 0
            },
            "behavior": "MOVE_AND_RESIZE"
          },
          "allowComments": false,
          "signingInstructions": "Review the budget before signing.",
          "suggestedSigner": "Ada Lovelace",
          "suggestedSigner2": "Finance",
          "suggestedSignerEmail": "ada@example.com",
          "caption": null,
          "invalidStamp": "invalid",
          "plainSignature": {
            "format": "PNG",
            "source": {
              "type": "INLINE_BASE64",
              "base64Data": "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII="
            }
          }
        }
      }
    },
    {
      "stepId": "read-signature",
      "target": {
        "type": "DRAWING_OBJECT_ALL_ON_SHEET",
        "sheetName": "Approvals"
      },
      "query": {
        "type": "GET_DRAWING_OBJECTS"
      }
    }
  ]
}
JSON

cat > "${streaming_request_path}" <<JSON
{
  "source": {
    "type": "NEW"
  },
  "persistence": {
    "type": "SAVE_AS",
    "path": "${streaming_workbook_rel}"
  },
  "execution": {
    "mode": {
      "writeMode": "STREAMING_WRITE"
    }
  },
  "steps": [
    {
      "stepId": "ensure-ledger",
      "target": {
        "type": "SHEET_BY_NAME",
        "name": "Ledger"
      },
      "action": {
        "type": "ENSURE_SHEET"
      }
    },
    {
      "stepId": "append-header",
      "target": {
        "type": "SHEET_BY_NAME",
        "name": "Ledger"
      },
      "action": {
        "type": "APPEND_ROW",
        "values": [
          {
            "type": "TEXT",
            "source": {
              "type": "INLINE",
              "text": "Team"
            }
          },
          {
            "type": "TEXT",
            "source": {
              "type": "INLINE",
              "text": "Task"
            }
          },
          {
            "type": "TEXT",
            "source": {
              "type": "INLINE",
              "text": "Hours"
            }
          }
        ]
      }
    },
    {
      "stepId": "append-ops",
      "target": {
        "type": "SHEET_BY_NAME",
        "name": "Ledger"
      },
      "action": {
        "type": "APPEND_ROW",
        "values": [
          {
            "type": "TEXT",
            "source": {
              "type": "INLINE",
              "text": "Ops"
            }
          },
          {
            "type": "TEXT",
            "source": {
              "type": "INLINE",
              "text": "Badge prep"
            }
          },
          {
            "type": "NUMBER",
            "number": 6.5
          }
        ]
      }
    }
  ]
}
JSON

cat > "${streaming_read_request_path}" <<JSON
{
  "source": {
    "type": "EXISTING",
    "path": "${streaming_workbook_rel}"
  },
  "persistence": {
    "type": "NONE"
  },
  "execution": {
    "mode": {
      "readMode": "EVENT_READ"
    }
  },
  "steps": [
    {
      "stepId": "streaming-workbook",
      "target": {
        "type": "WORKBOOK_CURRENT"
      },
      "query": {
        "type": "GET_WORKBOOK_SUMMARY"
      }
    },
    {
      "stepId": "streaming-sheet",
      "target": {
        "type": "SHEET_BY_NAME",
        "name": "Ledger"
      },
      "query": {
        "type": "GET_SHEET_SUMMARY"
      }
    }
  ]
}
JSON

printf 'Docker smoke: building local image\n'
docker_with_repo_config buildx build --load -t "${image_tag}" "${repo_root}" >/dev/null

printf 'Docker smoke: verifying packaged help and catalog contract\n'
if [[ -n "${docker_endpoint}" ]]; then
    DOCKER_CONFIG="${anonymous_docker_config}" DOCKER_HOST="${docker_endpoint}" \
        "${repo_root}/scripts/verify-cli-contract.sh" docker-image "${image_tag}" >/dev/null
else
    DOCKER_CONFIG="${anonymous_docker_config}" \
        "${repo_root}/scripts/verify-cli-contract.sh" docker-image "${image_tag}" >/dev/null
fi

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
[[ ! -f "${legacy_workbook_path}" ]] || die \
    "docker smoke wrote the workbook relative to the shell workdir instead of the request file directory: ${legacy_workbook_path}"
grep -Eq '"status"[[:space:]]*:[[:space:]]*"SUCCESS"' "${response_path}" || die \
    "docker smoke response did not report SUCCESS"
grep -Eq '"journal"[[:space:]]*:' "${response_path}" || die \
    "docker smoke response did not include the structured execution journal"
grep -Eq '"level"[[:space:]]*:[[:space:]]*"VERBOSE"' "${response_path}" || die \
    "docker smoke response did not preserve the requested VERBOSE journal level"
[[ -s "${create_stderr_path}" ]] || die \
    "docker smoke create request did not stream live verbose journal events to stderr"
grep -Fq '[gridgrind]' "${create_stderr_path}" || die \
    "docker smoke create request stderr did not contain the CLI journal prefix"

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

printf 'Docker smoke: verifying signature-line authoring under container fonts\n'
docker_with_repo_config run --rm \
    --user "${docker_run_user}" \
    -w /workdir \
    -v "${smoke_root}:/workdir" \
    "${image_tag}" \
    --request "${signature_request_rel}" \
    --response "${signature_response_rel}" >/dev/null 2>"${signature_stderr_path}"

[[ -f "${signature_response_path}" ]] || die \
    "docker smoke signature response file was not written: ${signature_response_path}"
[[ -f "${signature_workbook_path}" ]] || die \
    "docker smoke signature workbook file was not written: ${signature_workbook_path}"
[[ ! -f "${legacy_signature_workbook_path}" ]] || die \
    "docker smoke wrote the signature workbook relative to the shell workdir instead of the request file directory: ${legacy_signature_workbook_path}"
grep -Eq '"status"[[:space:]]*:[[:space:]]*"SUCCESS"' "${signature_response_path}" || die \
    "docker smoke signature-line authoring did not report SUCCESS"
grep -Eq '"BudgetSignature"' "${signature_response_path}" || die \
    "docker smoke signature-line response did not include the authored drawing object"
[[ ! -s "${signature_stderr_path}" ]] || die \
    "docker smoke signature-line request wrote unexpected stderr: $(tr '\n' ' ' < "${signature_stderr_path}")"

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
[[ ! -f "${legacy_streaming_workbook_path}" ]] || die \
    "docker smoke wrote the streaming workbook relative to the shell workdir instead of the request file directory: ${legacy_streaming_workbook_path}"
grep -Eq '"status"[[:space:]]*:[[:space:]]*"SUCCESS"' "${streaming_response_path}" || die \
    "docker smoke STREAMING_WRITE authoring did not report SUCCESS"
[[ ! -s "${streaming_stderr_path}" ]] || die \
    "docker smoke STREAMING_WRITE authoring wrote unexpected stderr: $(tr '\n' ' ' < "${streaming_stderr_path}")"

docker_with_repo_config run --rm \
    --user "${docker_run_user}" \
    -w /workdir \
    -v "${smoke_root}:/workdir" \
    "${image_tag}" \
    --request "${streaming_read_request_rel}" \
    --response "${streaming_read_response_rel}" >/dev/null 2>"${streaming_read_stderr_path}"

[[ -f "${streaming_read_response_path}" ]] || die \
    "docker smoke streaming readback response file was not written: ${streaming_read_response_path}"
grep -Eq '"status"[[:space:]]*:[[:space:]]*"SUCCESS"' "${streaming_read_response_path}" || die \
    "docker smoke STREAMING_WRITE readback did not report SUCCESS"
[[ ! -s "${streaming_read_stderr_path}" ]] || die \
    "docker smoke STREAMING_WRITE readback wrote unexpected stderr: $(tr '\n' ' ' < "${streaming_read_stderr_path}")"

printf 'Docker smoke: success\n'
