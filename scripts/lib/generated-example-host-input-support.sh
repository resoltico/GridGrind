#!/usr/bin/env bash
# Explicit host inputs for generated V3 fixture execution.

prepare_generated_example_host_inputs() {
    local fixture_grant_writer=$1
    local scratch_root=$2
    local package_security_request=$3
    local task_starter_request=$4
    local package_security_inspection_request=$5
    local secrets_provider_root="${scratch_root}/secrets-provider"

    mkdir -p "${secrets_provider_root}"
    printf '%s' 'GridGrind-2026' > "${secrets_provider_root}/output-password"
    printf '%s' 'GridGrind-2026' > "${secrets_provider_root}/source-open-password"

    "${fixture_grant_writer}" \
        --request "${package_security_request}" \
        --output "${scratch_root}/package-security-asset-grant.json"
    if [[ -f "${task_starter_request}" ]]; then
        "${fixture_grant_writer}" \
            --request "${task_starter_request}" \
            --output "${scratch_root}/task-starter-workbook-grant.json"
    fi
    if [[ -f "${package_security_inspection_request}" ]]; then
        "${fixture_grant_writer}" \
            --request "${package_security_inspection_request}" \
            --output "${scratch_root}/package-security-inspect-grant.json"
    fi
}
