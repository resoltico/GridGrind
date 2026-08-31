#!/usr/bin/env bash
# Runtime-aware verification for the documented bind-mounted Docker command surface.

# shellcheck source=/dev/null
source "$(cd -P -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)/docker-smoke-legal-support.sh"

verify_documented_bind_mount_user_guidance() {
    local probe_image_tag=$1
    local probe_smoke_root=$2
    local probe_docker_run_user=$3
    local probe_request_dir="${probe_smoke_root}/requests odd"
    local documented_no_user_request_rel='requests odd/request documented mount no-user [docker #smoke].json'
    local documented_no_user_response_rel='responses odd/nested/response documented mount no-user [docker #smoke].json'
    local documented_no_user_workbook_rel='books odd/nested/office documented mount no-user [docker #smoke].xlsx'
    local documented_with_user_request_rel='requests odd/request documented mount with-user [docker #smoke].json'
    local documented_with_user_response_rel='responses odd/nested/response documented mount with-user [docker #smoke].json'
    local documented_with_user_workbook_rel='books odd/nested/office documented mount with-user [docker #smoke].xlsx'
    local documented_no_user_request_path="${probe_smoke_root}/${documented_no_user_request_rel}"
    local documented_no_user_response_path="${probe_smoke_root}/${documented_no_user_response_rel}"
    local documented_no_user_legacy_workbook_path="${probe_smoke_root}/${documented_no_user_workbook_rel}"
    local documented_no_user_workbook_path="${probe_request_dir}/${documented_no_user_workbook_rel}"
    local documented_with_user_request_path="${probe_smoke_root}/${documented_with_user_request_rel}"
    local documented_with_user_response_path="${probe_smoke_root}/${documented_with_user_response_rel}"
    local documented_with_user_legacy_workbook_path="${probe_smoke_root}/${documented_with_user_workbook_rel}"
    local documented_with_user_workbook_path="${probe_request_dir}/${documented_with_user_workbook_rel}"
    local documented_no_user_stdout_path="${probe_smoke_root}/stdout documented no-user [docker #smoke].log"
    local documented_no_user_stderr_path="${probe_smoke_root}/stderr documented no-user [docker #smoke].log"
    local documented_with_user_stderr_path="${probe_smoke_root}/stderr documented with-user [docker #smoke].log"
    local documented_no_user_exit_code=0

    write_documented_request() {
        local target_request_path=$1
        local target_workbook_rel=$2

        cat > "${target_request_path}" <<JSON
{
  "protocolVersion": "V2",
  "source": {
    "type": "NEW"
  },
  "persistence": {
    "type": "SAVE_AS",
    "path": "${target_workbook_rel}",
    "ifExists": "REPLACE"
  },
  "execution": {
    "mode": {"type": "FULL_XSSF"},
    "journal": {
      "level": "NORMAL"
    },
    "calculation": {
      "strategy": {
        "type": "DO_NOT_CALCULATE"
      },
      "markRecalculateOnOpen": false
    }
  },
  "formulaEnvironment": {
    "externalWorkbooks": [],
    "missingWorkbookPolicy": "ERROR",
    "udfToolpacks": []
  },
  "steps": [
    {
      "stepId": "ensure-documented-mount",
      "target": {
        "type": "SHEET_BY_NAME",
        "name": "DocumentedMount"
      },
      "action": {
        "type": "ENSURE_SHEET"
      }
    },
    {
      "stepId": "documented-workbook",
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
    }

    write_documented_request "${documented_no_user_request_path}" "${documented_no_user_workbook_rel}"
    write_documented_request "${documented_with_user_request_path}" "${documented_with_user_workbook_rel}"
    mkdir -p \
        "$(dirname -- "${documented_no_user_response_path}")" \
        "$(dirname -- "${documented_no_user_workbook_path}")" \
        "$(dirname -- "${documented_with_user_response_path}")" \
        "$(dirname -- "${documented_with_user_workbook_path}")"

    rm -f \
        "${documented_no_user_response_path}" \
        "${documented_no_user_workbook_path}" \
        "${documented_no_user_legacy_workbook_path}" \
        "${documented_with_user_response_path}" \
        "${documented_with_user_workbook_path}" \
        "${documented_with_user_legacy_workbook_path}" \
        "${documented_no_user_stdout_path}" \
        "${documented_no_user_stderr_path}" \
        "${documented_with_user_stderr_path}"
    set +e
    docker_with_repo_config run --rm -i \
        -v "${probe_smoke_root}:/work" \
        "${probe_image_tag}" \
        --request "${documented_no_user_request_rel}" \
        --response "${documented_no_user_response_rel}" >"${documented_no_user_stdout_path}" \
        2>"${documented_no_user_stderr_path}"
    documented_no_user_exit_code=$?
    set -e

    if [[ ${documented_no_user_exit_code} -eq 0 ]]; then
        [[ -f "${documented_no_user_response_path}" ]] || die \
            "docker smoke no-user documented bind-mount run did not write the response file"
        [[ -f "${documented_no_user_workbook_path}" ]] || die \
            "docker smoke no-user documented bind-mount run did not write the workbook file"
        [[ ! -f "${documented_no_user_legacy_workbook_path}" ]] || die \
            "docker smoke no-user documented bind-mount run wrote the workbook relative to the shell workdir"
        [[ ! -s "${documented_no_user_stdout_path}" ]] || die \
            "docker smoke no-user documented bind-mount run wrote unexpected stdout on a remapped runtime"
        [[ ! -s "${documented_no_user_stderr_path}" ]] || die \
            "docker smoke no-user documented bind-mount run wrote unexpected stderr on a remapped runtime"
    else
        [[ ! -f "${documented_no_user_legacy_workbook_path}" ]] || die \
            "docker smoke no-user documented bind-mount run wrote the workbook relative to the shell workdir"
        if [[ -f "${documented_no_user_response_path}" && -f "${documented_no_user_workbook_path}" ]]; then
            die \
                "docker smoke no-user documented bind-mount run completed both mounted writes despite failing"
        fi
    fi

    docker_with_repo_config run --rm -i \
        --user "${probe_docker_run_user}" \
        -v "${probe_smoke_root}:/work" \
        "${probe_image_tag}" \
        --request "${documented_with_user_request_rel}" \
        --response "${documented_with_user_response_rel}" >/dev/null 2>"${documented_with_user_stderr_path}"
    [[ -f "${documented_with_user_response_path}" ]] || die \
        "docker smoke documented bind-mount run with --user did not write the response file"
    [[ -f "${documented_with_user_workbook_path}" ]] || die \
        "docker smoke documented bind-mount run with --user did not write the workbook file"
    [[ ! -f "${documented_with_user_legacy_workbook_path}" ]] || die \
        "docker smoke documented bind-mount run with --user wrote the workbook relative to the shell workdir"
    grep -Eq '"status"[[:space:]]*:[[:space:]]*"SUCCEEDED"' "${documented_with_user_response_path}" || die \
        "docker smoke documented bind-mount run with --user did not report SUCCEEDED"
    [[ ! -s "${documented_with_user_stderr_path}" ]] || die \
        "docker smoke documented bind-mount run with --user wrote unexpected stderr: $(tr '\n' ' ' < "${documented_with_user_stderr_path}")"
    rm -f \
        "${documented_no_user_request_path}" \
        "${documented_no_user_response_path}" \
        "${documented_no_user_workbook_path}" \
        "${documented_no_user_legacy_workbook_path}" \
        "${documented_with_user_request_path}" \
        "${documented_with_user_response_path}" \
        "${documented_with_user_workbook_path}" \
        "${documented_with_user_legacy_workbook_path}" \
        "${documented_no_user_stdout_path}" \
        "${documented_no_user_stderr_path}" \
        "${documented_with_user_stderr_path}"
}
