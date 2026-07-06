#!/usr/bin/env bash
# Shared fake-Docker and fixture-runner support for the container-publication regression script.

write_fake_docker_cli() {
    local fake_docker_path=$1
    cat > "${fake_docker_path}" <<'EOF'
#!/usr/bin/env bash
set -euo pipefail

emit_fixture_file() {
    local fixture_path=$1
    [[ -f "${fixture_path}" ]] || {
        printf 'missing fixture file: %s\n' "${fixture_path}" >&2
        exit 1
    }
    cat "${fixture_path}"
}

log_path=${FAKE_DOCKER_LOG:?}
printf '%s\n' "$*" >> "${log_path}"

args=("$@")
offset=0
if [[ ${#args[@]} -ge 2 && "${args[0]}" == "--config" ]]; then
    offset=2
fi

if [[ ${#args[@]} -le ${offset} ]]; then
    printf 'unexpected docker invocation: %s\n' "$*" >&2
    exit 1
fi

command=${args[${offset}]}
case "${command}" in
    pull)
        image_ref=${args[$((offset + 1))]:-}
        [[ -n "${image_ref}" ]] || exit 1
        exit 0
        ;;
    run)
        run_index=$((offset + 1))
        while [[ ${#args[@]} -gt ${run_index} ]]; do
            case "${args[${run_index}]}" in
                --rm|--interactive|--tty|-i|-t|-it|-ti)
                    ((run_index += 1))
                    ;;
                *)
                    break
                    ;;
            esac
        done
        image_ref=${args[${run_index}]:-}
        cli_flag=${args[$((run_index + 1))]:-}
        case "${cli_flag}" in
            '')
                emit_fixture_file "${FAKE_DOCKER_NOARGS_FAILURE_OUTPUT_FILE:?}" >&2
                exit 2
                ;;
            --version)
                case "${image_ref}" in
                    *:latest)
                        emit_fixture_file "${FAKE_DOCKER_LATEST_VERSION_OUTPUT_FILE:?}"
                        ;;
                    *)
                        emit_fixture_file "${FAKE_DOCKER_VERSION_OUTPUT_FILE:?}"
                        ;;
                esac
                ;;
            --help)
                emit_fixture_file "${FAKE_DOCKER_HELP_OVERVIEW_OUTPUT_FILE:?}"
                ;;
            --help-protocol)
                emit_fixture_file "${FAKE_DOCKER_HELP_PROTOCOL_OUTPUT_FILE:?}"
                ;;
            --help-guidance)
                emit_fixture_file "${FAKE_DOCKER_HELP_GUIDANCE_OUTPUT_FILE:?}"
                ;;
            --print-request-template)
                emit_fixture_file "${FAKE_DOCKER_REQUEST_TEMPLATE_OUTPUT_FILE:?}"
                ;;
            --print-recipe-catalog)
                if [[ "${args[$((run_index + 2))]:-}" == "--lookup" ]]; then
                    case "${args[$((run_index + 3))]:-}" in
                        BUDGET)
                            emit_fixture_file "${FAKE_DOCKER_EXAMPLE_RECIPE_CATALOG_DETAIL_OUTPUT_FILE:?}"
                            ;;
                        TABULAR_REPORT)
                            emit_fixture_file "${FAKE_DOCKER_TASK_RECIPE_CATALOG_DETAIL_OUTPUT_FILE:?}"
                            ;;
                        *)
                            printf 'unexpected docker recipe catalog lookup fixture request: %s\n' "${args[$((run_index + 3))]:-}" >&2
                            exit 1
                            ;;
                    esac
                else
                    emit_fixture_file "${FAKE_DOCKER_RECIPE_CATALOG_OUTPUT_FILE:?}"
                fi
                ;;
            --print-recipe)
                emit_fixture_file "${FAKE_DOCKER_RECIPE_REQUEST_OUTPUT_FILE:?}"
                ;;
            --print-recipe-keyword-match)
                emit_fixture_file "${FAKE_DOCKER_RECIPE_KEYWORD_MATCH_OUTPUT_FILE:?}"
                ;;
            --doctor-request)
                emit_fixture_file "${FAKE_DOCKER_DOCTOR_REPORT_OUTPUT_FILE:?}"
                ;;
            --print-protocol-catalog)
                if [[ "${args[$((run_index + 2))]:-}" == "--lookup" ]]; then
                    case "${args[$((run_index + 3))]:-}" in
                        sourceTypes)
                            emit_fixture_file "${FAKE_DOCKER_SOURCE_TYPES_OUTPUT_FILE:?}"
                            ;;
                        persistenceTypes)
                            emit_fixture_file "${FAKE_DOCKER_PERSISTENCE_TYPES_OUTPUT_FILE:?}"
                            ;;
                        stepTypes)
                            emit_fixture_file "${FAKE_DOCKER_STEP_TYPES_OUTPUT_FILE:?}"
                            ;;
                        mutationActionTypes)
                            emit_fixture_file "${FAKE_DOCKER_MUTATION_ACTION_TYPES_OUTPUT_FILE:?}"
                            ;;
                        assertionTypes)
                            emit_fixture_file "${FAKE_DOCKER_ASSERTION_TYPES_OUTPUT_FILE:?}"
                            ;;
                        inspectionQueryTypes)
                            emit_fixture_file "${FAKE_DOCKER_INSPECTION_QUERY_TYPES_OUTPUT_FILE:?}"
                            ;;
                        nestedTypes:executionModeTypes)
                            emit_fixture_file "${FAKE_DOCKER_EXECUTION_MODE_TYPES_OUTPUT_FILE:?}"
                            ;;
                        plainTypes:executionPolicyInputType)
                            emit_fixture_file "${FAKE_DOCKER_EXECUTION_POLICY_INPUT_TYPE_OUTPUT_FILE:?}"
                            ;;
                        *)
                            printf 'unexpected docker protocol lookup fixture request: %s\n' "${args[$((run_index + 3))]:-}" >&2
                            exit 1
                            ;;
                    esac
                else
                    emit_fixture_file "${FAKE_DOCKER_CATALOG_INDEX_OUTPUT_FILE:?}"
                fi
                ;;
            *)
                printf 'unexpected docker run invocation: %s\n' "$*" >&2
                exit 1
                ;;
        esac
        ;;
    *)
        printf 'unexpected docker subcommand: %s\n' "${command}" >&2
        exit 1
        ;;
esac
EOF
    chmod +x "${fake_docker_path}"
}

write_fake_discovery_verify_script() {
    local fake_discovery_path=$1
    cat > "${fake_discovery_path}" <<'EOF'
#!/usr/bin/env bash
set -euo pipefail

printf 'discovery %s\n' "$*" >> "${FAKE_DOCKER_LOG:?}"
if [[ "${FAKE_DISCOVERY_SHOULD_FAIL:-0}" == 1 ]]; then
    printf 'forced discovery failure for %s\n' "$*" >&2
    exit 1
fi
exit 0
EOF
    chmod +x "${fake_discovery_path}"
}

run_fake_docker_verify_with_fixture_texts() {
    local version_output=$1
    local latest_version_output=$2
    local help_overview_output=$3
    local help_protocol_output=$4
    local help_guidance_output=$5
    local catalog_index_output=$6
    local source_types_output=$7
    local persistence_types_output=$8
    local step_types_output=$9
    local mutation_action_types_output=${10}
    local assertion_types_output=${11}
    local inspection_query_types_output=${12}
    local execution_mode_types_output=${13}
    local execution_policy_input_type_output=${14}
    local recipe_catalog_output=${15}
    local recipe_request_output=${16}
    local recipe_keyword_match_output=${17}
    local doctor_report_output=${18}
    local noargs_failure_output=${19}
    local example_recipe_catalog_detail_output=${20:-${success_example_recipe_catalog_detail}}
    local task_recipe_catalog_detail_output=${21:-${success_task_recipe_catalog_detail}}
    local case_dir version_output_file latest_version_output_file
    local help_overview_output_file help_protocol_output_file help_guidance_output_file
    local catalog_index_output_file source_types_output_file persistence_types_output_file
    local step_types_output_file mutation_action_types_output_file assertion_types_output_file
    local inspection_query_types_output_file execution_mode_types_output_file
    local execution_policy_input_type_output_file recipe_catalog_output_file
    local example_recipe_catalog_detail_output_file task_recipe_catalog_detail_output_file
    local recipe_request_output_file recipe_keyword_match_output_file
    local request_template_output_file doctor_report_output_file noargs_failure_output_file

    case_dir="$(next_fixture_case_dir "${test_root}")"
    version_output_file="$(write_case_fixture "${case_dir}" 'version.txt' "${version_output}")"
    latest_version_output_file="$(write_case_fixture "${case_dir}" 'latest-version.txt' "${latest_version_output}")"
    help_overview_output_file="$(write_case_fixture "${case_dir}" 'help-overview.txt' "${help_overview_output}")"
    help_protocol_output_file="$(write_case_fixture "${case_dir}" 'help-protocol.txt' "${help_protocol_output}")"
    help_guidance_output_file="$(write_case_fixture "${case_dir}" 'help-guidance.txt' "${help_guidance_output}")"
    catalog_index_output_file="$(write_case_fixture "${case_dir}" 'protocol-catalog-index.json' "${catalog_index_output}")"
    source_types_output_file="$(write_case_fixture "${case_dir}" 'source-types.json' "${source_types_output}")"
    persistence_types_output_file="$(write_case_fixture "${case_dir}" 'persistence-types.json' "${persistence_types_output}")"
    step_types_output_file="$(write_case_fixture "${case_dir}" 'step-types.json' "${step_types_output}")"
    mutation_action_types_output_file="$(write_case_fixture "${case_dir}" 'mutation-action-types.json' "${mutation_action_types_output}")"
    assertion_types_output_file="$(write_case_fixture "${case_dir}" 'assertion-types.json' "${assertion_types_output}")"
    inspection_query_types_output_file="$(write_case_fixture "${case_dir}" 'inspection-query-types.json' "${inspection_query_types_output}")"
    execution_mode_types_output_file="$(write_case_fixture "${case_dir}" 'execution-mode-types.json' "${execution_mode_types_output}")"
    execution_policy_input_type_output_file="$(write_case_fixture "${case_dir}" 'execution-policy-input-type.json' "${execution_policy_input_type_output}")"
    recipe_catalog_output_file="$(write_case_fixture "${case_dir}" 'recipe-catalog.json' "${recipe_catalog_output}")"
    example_recipe_catalog_detail_output_file="$(write_case_fixture "${case_dir}" 'recipe-catalog-example-detail.json' "${example_recipe_catalog_detail_output}")"
    task_recipe_catalog_detail_output_file="$(write_case_fixture "${case_dir}" 'recipe-catalog-task-detail.json' "${task_recipe_catalog_detail_output}")"
    recipe_request_output_file="$(write_case_fixture "${case_dir}" 'recipe-request.json' "${recipe_request_output}")"
    recipe_keyword_match_output_file="$(write_case_fixture "${case_dir}" 'recipe-keyword-match.json' "${recipe_keyword_match_output}")"
    request_template_output_file="$(write_case_fixture "${case_dir}" 'request-template.json' "${success_request_template}")"
    doctor_report_output_file="$(write_case_fixture "${case_dir}" 'doctor-report.json' "${doctor_report_output}")"
    noargs_failure_output_file="$(write_case_fixture "${case_dir}" 'noargs-failure.json' "${noargs_failure_output}")"

    PATH="${fake_bin}:${PATH}" \
        FAKE_DOCKER_LOG="${fake_log}" \
        GRIDGRIND_VERIFY_CLI_DISCOVERY_EXECUTION_SCRIPT="${fake_discovery_verify}" \
        FAKE_DOCKER_VERSION_OUTPUT_FILE="${version_output_file}" \
        FAKE_DOCKER_LATEST_VERSION_OUTPUT_FILE="${latest_version_output_file}" \
        FAKE_DOCKER_HELP_OVERVIEW_OUTPUT_FILE="${help_overview_output_file}" \
        FAKE_DOCKER_HELP_PROTOCOL_OUTPUT_FILE="${help_protocol_output_file}" \
        FAKE_DOCKER_HELP_GUIDANCE_OUTPUT_FILE="${help_guidance_output_file}" \
        FAKE_DOCKER_CATALOG_INDEX_OUTPUT_FILE="${catalog_index_output_file}" \
        FAKE_DOCKER_SOURCE_TYPES_OUTPUT_FILE="${source_types_output_file}" \
        FAKE_DOCKER_PERSISTENCE_TYPES_OUTPUT_FILE="${persistence_types_output_file}" \
        FAKE_DOCKER_STEP_TYPES_OUTPUT_FILE="${step_types_output_file}" \
        FAKE_DOCKER_MUTATION_ACTION_TYPES_OUTPUT_FILE="${mutation_action_types_output_file}" \
        FAKE_DOCKER_ASSERTION_TYPES_OUTPUT_FILE="${assertion_types_output_file}" \
        FAKE_DOCKER_INSPECTION_QUERY_TYPES_OUTPUT_FILE="${inspection_query_types_output_file}" \
        FAKE_DOCKER_EXECUTION_MODE_TYPES_OUTPUT_FILE="${execution_mode_types_output_file}" \
        FAKE_DOCKER_EXECUTION_POLICY_INPUT_TYPE_OUTPUT_FILE="${execution_policy_input_type_output_file}" \
        FAKE_DOCKER_RECIPE_CATALOG_OUTPUT_FILE="${recipe_catalog_output_file}" \
        FAKE_DOCKER_EXAMPLE_RECIPE_CATALOG_DETAIL_OUTPUT_FILE="${example_recipe_catalog_detail_output_file}" \
        FAKE_DOCKER_TASK_RECIPE_CATALOG_DETAIL_OUTPUT_FILE="${task_recipe_catalog_detail_output_file}" \
        FAKE_DOCKER_RECIPE_REQUEST_OUTPUT_FILE="${recipe_request_output_file}" \
        FAKE_DOCKER_RECIPE_KEYWORD_MATCH_OUTPUT_FILE="${recipe_keyword_match_output_file}" \
        FAKE_DOCKER_REQUEST_TEMPLATE_OUTPUT_FILE="${request_template_output_file}" \
        FAKE_DOCKER_DOCTOR_REPORT_OUTPUT_FILE="${doctor_report_output_file}" \
        FAKE_DOCKER_NOARGS_FAILURE_OUTPUT_FILE="${noargs_failure_output_file}" \
        GRIDGRIND_PUBLICATION_VERIFY_RETRIES=1 \
        GRIDGRIND_PUBLICATION_VERIFY_DELAY_SECONDS=0 \
        "${verify_script}" "ghcr.io/example/gridgrind" "9.9.9" >/dev/null
}
