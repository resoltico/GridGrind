#!/usr/bin/env bash
# Shared fake-CLI and fixture-runner support for the packaged CLI contract regression script.

write_fake_gridgrind_cli() {
    local fake_cli_path=$1
    cat > "${fake_cli_path}" <<'EOF'
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

case "${1:-}" in
    '')
        if [[ -t 0 || -t 1 || -t 2 ]]; then
            emit_fixture_file "${FAKE_GRIDGRIND_INTERACTIVE_NOARGS_FAILURE_FILE:-${FAKE_GRIDGRIND_NOARGS_FAILURE_FILE:?}}"
        else
            emit_fixture_file "${FAKE_GRIDGRIND_NOARGS_FAILURE_FILE:?}" >&2
        fi
        exit 2
        ;;
    --help)
        emit_fixture_file "${FAKE_GRIDGRIND_HELP_OVERVIEW_FILE:?}"
        ;;
    --help-protocol)
        emit_fixture_file "${FAKE_GRIDGRIND_HELP_PROTOCOL_FILE:?}"
        ;;
    --help-guidance)
        emit_fixture_file "${FAKE_GRIDGRIND_HELP_GUIDANCE_FILE:?}"
        ;;
    --print-request-template)
        emit_fixture_file "${FAKE_GRIDGRIND_REQUEST_TEMPLATE_FILE:?}"
        ;;
    --doctor-request)
        emit_fixture_file "${FAKE_GRIDGRIND_DOCTOR_REPORT_FILE:?}"
        ;;
    --print-recipe-keyword-match)
        emit_fixture_file "${FAKE_GRIDGRIND_RECIPE_KEYWORD_MATCH_FILE:?}"
        ;;
    --print-recipe-catalog)
        if [[ "${2:-}" == '--lookup' ]]; then
            case "${3:-}" in
                BUDGET)
                    emit_fixture_file "${FAKE_GRIDGRIND_EXAMPLE_RECIPE_CATALOG_DETAIL_FILE:?}"
                    ;;
                TABULAR_REPORT)
                    emit_fixture_file "${FAKE_GRIDGRIND_TASK_RECIPE_CATALOG_DETAIL_FILE:?}"
                    ;;
                *)
                    printf 'unexpected recipe catalog lookup fixture request: %s\n' "${3:-}" >&2
                    exit 1
                    ;;
            esac
        else
            emit_fixture_file "${FAKE_GRIDGRIND_RECIPE_CATALOG_FILE:?}"
        fi
        ;;
    --print-recipe)
        emit_fixture_file "${FAKE_GRIDGRIND_RECIPE_REQUEST_FILE:?}"
        ;;
    --print-protocol-catalog)
        if [[ "${2:-}" == '--lookup' ]]; then
            case "${3:-}" in
                sourceTypes)
                    emit_fixture_file "${FAKE_GRIDGRIND_SOURCE_TYPES_FILE:?}"
                    ;;
                persistenceTypes)
                    emit_fixture_file "${FAKE_GRIDGRIND_PERSISTENCE_TYPES_FILE:?}"
                    ;;
                stepTypes)
                    emit_fixture_file "${FAKE_GRIDGRIND_STEP_TYPES_FILE:?}"
                    ;;
                mutationActionTypes)
                    emit_fixture_file "${FAKE_GRIDGRIND_MUTATION_ACTION_TYPES_FILE:?}"
                    ;;
                assertionTypes)
                    emit_fixture_file "${FAKE_GRIDGRIND_ASSERTION_TYPES_FILE:?}"
                    ;;
                inspectionQueryTypes)
                    emit_fixture_file "${FAKE_GRIDGRIND_INSPECTION_QUERY_TYPES_FILE:?}"
                    ;;
                nestedTypes:executionModeTypes)
                    emit_fixture_file "${FAKE_GRIDGRIND_EXECUTION_MODE_TYPES_FILE:?}"
                    ;;
                plainTypes:executionPolicyInputType)
                    emit_fixture_file "${FAKE_GRIDGRIND_EXECUTION_POLICY_INPUT_TYPE_FILE:?}"
                    ;;
                *)
                    printf 'unexpected protocol catalog lookup fixture request: %s\n' "${3:-}" >&2
                    exit 1
                    ;;
            esac
        else
            emit_fixture_file "${FAKE_GRIDGRIND_CATALOG_INDEX_FILE:?}"
        fi
        ;;
    *)
        printf 'unexpected invocation: %s\n' "$*" >&2
        exit 1
        ;;
esac
EOF
    chmod +x "${fake_cli_path}"
}

run_fake_gridgrind_verify_with_fixture_texts() {
    local unset_tmpdir=$1
    local help_overview_text=$2
    local help_protocol_text=$3
    local help_guidance_text=$4
    local source_types_text=$5
    local persistence_types_text=$6
    local step_types_text=$7
    local mutation_action_types_text=$8
    local assertion_types_text=$9
    local inspection_query_types_text=${10}
    local execution_mode_types_text=${11}
    local execution_policy_input_type_text=${12}
    local recipe_catalog_text=${13}
    local recipe_request_text=${14}
    local recipe_keyword_match_text=${15}
    local doctor_report_text=${16}
    local noargs_failure_text=${17}
    local interactive_noargs_failure_text=${18}
    local example_recipe_catalog_detail_text=${19:-${success_example_recipe_catalog_detail}}
    local task_recipe_catalog_detail_text=${20:-${success_task_recipe_catalog_detail}}
    local case_dir
    local help_overview_file
    local help_protocol_file
    local help_guidance_file
    local catalog_index_file
    local source_types_file
    local persistence_types_file
    local step_types_file
    local mutation_action_types_file
    local assertion_types_file
    local inspection_query_types_file
    local execution_mode_types_file
    local execution_policy_input_type_file
    local recipe_catalog_file
    local example_recipe_catalog_detail_file
    local task_recipe_catalog_detail_file
    local recipe_request_file
    local recipe_keyword_match_file
    local request_template_file
    local doctor_report_file
    local noargs_failure_file
    local interactive_noargs_failure_file

    case_dir="$(next_fixture_case_dir "${test_root}")"
    help_overview_file="$(write_case_fixture "${case_dir}" 'help-overview.txt' "${help_overview_text}")"
    help_protocol_file="$(write_case_fixture "${case_dir}" 'help-protocol.txt' "${help_protocol_text}")"
    help_guidance_file="$(write_case_fixture "${case_dir}" 'help-guidance.txt' "${help_guidance_text}")"
    catalog_index_file="$(write_case_fixture "${case_dir}" 'protocol-catalog-index.json' "${success_catalog_index}")"
    source_types_file="$(write_case_fixture "${case_dir}" 'source-types.json' "${source_types_text}")"
    persistence_types_file="$(write_case_fixture "${case_dir}" 'persistence-types.json' "${persistence_types_text}")"
    step_types_file="$(write_case_fixture "${case_dir}" 'step-types.json' "${step_types_text}")"
    mutation_action_types_file="$(write_case_fixture "${case_dir}" 'mutation-action-types.json' "${mutation_action_types_text}")"
    assertion_types_file="$(write_case_fixture "${case_dir}" 'assertion-types.json' "${assertion_types_text}")"
    inspection_query_types_file="$(write_case_fixture "${case_dir}" 'inspection-query-types.json' "${inspection_query_types_text}")"
    execution_mode_types_file="$(write_case_fixture "${case_dir}" 'execution-mode-types.json' "${execution_mode_types_text}")"
    execution_policy_input_type_file="$(write_case_fixture "${case_dir}" 'execution-policy-input-type.json' "${execution_policy_input_type_text}")"
    recipe_catalog_file="$(write_case_fixture "${case_dir}" 'recipe-catalog.json' "${recipe_catalog_text}")"
    example_recipe_catalog_detail_file="$(write_case_fixture "${case_dir}" 'recipe-catalog-example-detail.json' "${example_recipe_catalog_detail_text}")"
    task_recipe_catalog_detail_file="$(write_case_fixture "${case_dir}" 'recipe-catalog-task-detail.json' "${task_recipe_catalog_detail_text}")"
    recipe_request_file="$(write_case_fixture "${case_dir}" 'recipe-request.json' "${recipe_request_text}")"
    recipe_keyword_match_file="$(write_case_fixture "${case_dir}" 'recipe-keyword-match.json' "${recipe_keyword_match_text}")"
    request_template_file="$(write_case_fixture "${case_dir}" 'request-template.json' "${success_request_template}")"
    doctor_report_file="$(write_case_fixture "${case_dir}" 'doctor-report.json' "${doctor_report_text}")"
    noargs_failure_file="$(write_case_fixture "${case_dir}" 'noargs-failure.json' "${noargs_failure_text}")"
    interactive_noargs_failure_file="$(
        write_case_fixture "${case_dir}" 'interactive-noargs-failure.json' "${interactive_noargs_failure_text}"
    )"

    if [[ "${unset_tmpdir}" == 'true' ]]; then
        env -u TMPDIR \
            FAKE_GRIDGRIND_HELP_OVERVIEW_FILE="${help_overview_file}" \
            FAKE_GRIDGRIND_HELP_PROTOCOL_FILE="${help_protocol_file}" \
            FAKE_GRIDGRIND_HELP_GUIDANCE_FILE="${help_guidance_file}" \
            FAKE_GRIDGRIND_CATALOG_INDEX_FILE="${catalog_index_file}" \
            FAKE_GRIDGRIND_SOURCE_TYPES_FILE="${source_types_file}" \
            FAKE_GRIDGRIND_PERSISTENCE_TYPES_FILE="${persistence_types_file}" \
            FAKE_GRIDGRIND_STEP_TYPES_FILE="${step_types_file}" \
            FAKE_GRIDGRIND_MUTATION_ACTION_TYPES_FILE="${mutation_action_types_file}" \
            FAKE_GRIDGRIND_ASSERTION_TYPES_FILE="${assertion_types_file}" \
            FAKE_GRIDGRIND_INSPECTION_QUERY_TYPES_FILE="${inspection_query_types_file}" \
            FAKE_GRIDGRIND_EXECUTION_MODE_TYPES_FILE="${execution_mode_types_file}" \
            FAKE_GRIDGRIND_EXECUTION_POLICY_INPUT_TYPE_FILE="${execution_policy_input_type_file}" \
            FAKE_GRIDGRIND_RECIPE_CATALOG_FILE="${recipe_catalog_file}" \
            FAKE_GRIDGRIND_EXAMPLE_RECIPE_CATALOG_DETAIL_FILE="${example_recipe_catalog_detail_file}" \
            FAKE_GRIDGRIND_TASK_RECIPE_CATALOG_DETAIL_FILE="${task_recipe_catalog_detail_file}" \
            FAKE_GRIDGRIND_RECIPE_REQUEST_FILE="${recipe_request_file}" \
            FAKE_GRIDGRIND_RECIPE_KEYWORD_MATCH_FILE="${recipe_keyword_match_file}" \
            FAKE_GRIDGRIND_REQUEST_TEMPLATE_FILE="${request_template_file}" \
            FAKE_GRIDGRIND_DOCTOR_REPORT_FILE="${doctor_report_file}" \
            FAKE_GRIDGRIND_NOARGS_FAILURE_FILE="${noargs_failure_file}" \
            FAKE_GRIDGRIND_INTERACTIVE_NOARGS_FAILURE_FILE="${interactive_noargs_failure_file}" \
            "${verify_script}" binary "${fake_cli}" >/dev/null
        return 0
    fi

    FAKE_GRIDGRIND_HELP_OVERVIEW_FILE="${help_overview_file}" \
        FAKE_GRIDGRIND_HELP_PROTOCOL_FILE="${help_protocol_file}" \
        FAKE_GRIDGRIND_HELP_GUIDANCE_FILE="${help_guidance_file}" \
        FAKE_GRIDGRIND_CATALOG_INDEX_FILE="${catalog_index_file}" \
        FAKE_GRIDGRIND_SOURCE_TYPES_FILE="${source_types_file}" \
        FAKE_GRIDGRIND_PERSISTENCE_TYPES_FILE="${persistence_types_file}" \
        FAKE_GRIDGRIND_STEP_TYPES_FILE="${step_types_file}" \
        FAKE_GRIDGRIND_MUTATION_ACTION_TYPES_FILE="${mutation_action_types_file}" \
        FAKE_GRIDGRIND_ASSERTION_TYPES_FILE="${assertion_types_file}" \
        FAKE_GRIDGRIND_INSPECTION_QUERY_TYPES_FILE="${inspection_query_types_file}" \
        FAKE_GRIDGRIND_EXECUTION_MODE_TYPES_FILE="${execution_mode_types_file}" \
        FAKE_GRIDGRIND_EXECUTION_POLICY_INPUT_TYPE_FILE="${execution_policy_input_type_file}" \
        FAKE_GRIDGRIND_RECIPE_CATALOG_FILE="${recipe_catalog_file}" \
        FAKE_GRIDGRIND_EXAMPLE_RECIPE_CATALOG_DETAIL_FILE="${example_recipe_catalog_detail_file}" \
        FAKE_GRIDGRIND_TASK_RECIPE_CATALOG_DETAIL_FILE="${task_recipe_catalog_detail_file}" \
        FAKE_GRIDGRIND_RECIPE_REQUEST_FILE="${recipe_request_file}" \
        FAKE_GRIDGRIND_RECIPE_KEYWORD_MATCH_FILE="${recipe_keyword_match_file}" \
        FAKE_GRIDGRIND_REQUEST_TEMPLATE_FILE="${request_template_file}" \
        FAKE_GRIDGRIND_DOCTOR_REPORT_FILE="${doctor_report_file}" \
        FAKE_GRIDGRIND_NOARGS_FAILURE_FILE="${noargs_failure_file}" \
        FAKE_GRIDGRIND_INTERACTIVE_NOARGS_FAILURE_FILE="${interactive_noargs_failure_file}" \
        "${verify_script}" binary "${fake_cli}" >/dev/null
}
