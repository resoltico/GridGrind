#!/usr/bin/env bash
# Shared case-directory helpers for CLI contract and publication regression fixtures.

next_fixture_case_dir() {
    local case_root=$1
    verify_case_counter=$((verify_case_counter + 1))
    local case_dir="${case_root}/verify-case-${verify_case_counter}"
    mkdir -p "${case_dir}"
    printf '%s' "${case_dir}"
}

write_case_fixture() {
    local case_dir=$1
    local fixture_name=$2
    local fixture_text=$3
    local fixture_path="${case_dir}/${fixture_name}"

    printf '%s' "${fixture_text}" > "${fixture_path}"
    printf '%s' "${fixture_path}"
}
