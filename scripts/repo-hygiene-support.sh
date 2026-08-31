#!/usr/bin/env bash
# Shared root-boundary and local-state definitions for repository hygiene tooling.

readonly repo_hygiene_structural_root_entries=(
    .codex
    .devcontainer
    .dockerignore
    .git
    .gitattributes
    .github
    .gitignore
    AGENTS.md
    CHANGELOG.md
    Dockerfile
    LICENSE
    LICENSE-APACHE-2.0
    LICENSE-BSD-2-CLAUSE
    LICENSE-BSD-3-CLAUSE
    NOTICE
    PATENTS.md
    README.md
    authoring-java
    build.gradle.kts
    check.sh
    check_mutation.sh
    cli
    contract
    docs
    engine
    examples
    excel-foundation
    executor
    gradle
    gradle.properties
    gradlew
    gradlew.bat
    images
    jazzer
    scripts
    settings.gradle.kts
)

readonly repo_hygiene_local_state_root_entries=(
    .claude
    .gradle
    .gradle-home
    .vscode
    build
    buildSrc
    generated-workbooks
    tmp
)

readonly repo_hygiene_generated_state_root_entries=(
    .gradle
    .gradle-home
    build
    buildSrc
    generated-workbooks
)

readonly repo_hygiene_tool_state_root_entries=(
    .claude
    .vscode
)

readonly repo_hygiene_scratch_root_entries=(
    tmp
)

repo_hygiene_contains() {
    local candidate=$1
    shift
    local value
    for value in "$@"; do
        if [[ "${value}" == "${candidate}" ]]; then
            return 0
        fi
    done
    return 1
}

repo_hygiene_is_structural_root_entry() {
    repo_hygiene_contains "$1" "${repo_hygiene_structural_root_entries[@]}"
}

repo_hygiene_is_local_state_root_entry() {
    repo_hygiene_contains "$1" "${repo_hygiene_local_state_root_entries[@]}"
}

repo_hygiene_is_generated_state_root_entry() {
    repo_hygiene_contains "$1" "${repo_hygiene_generated_state_root_entries[@]}"
}

repo_hygiene_is_tool_state_root_entry() {
    repo_hygiene_contains "$1" "${repo_hygiene_tool_state_root_entries[@]}"
}

repo_hygiene_is_scratch_root_entry() {
    repo_hygiene_contains "$1" "${repo_hygiene_scratch_root_entries[@]}"
}

repo_hygiene_is_expected_root_entry() {
    local entry_name=$1
    repo_hygiene_is_structural_root_entry "${entry_name}" ||
        repo_hygiene_is_local_state_root_entry "${entry_name}"
}

repo_hygiene_local_state_category() {
    local entry_name=$1
    if repo_hygiene_is_generated_state_root_entry "${entry_name}"; then
        printf '%s\n' 'generated-state'
        return 0
    fi
    if repo_hygiene_is_tool_state_root_entry "${entry_name}"; then
        printf '%s\n' 'tool-state'
        return 0
    fi
    if repo_hygiene_is_scratch_root_entry "${entry_name}"; then
        printf '%s\n' 'scratch-state'
        return 0
    fi
    printf '%s\n' 'local-state'
}

repo_hygiene_local_state_cleanup_flag() {
    local entry_name=$1
    if repo_hygiene_is_generated_state_root_entry "${entry_name}"; then
        printf '%s\n' '--purge-generated-state'
        return 0
    fi
    if repo_hygiene_is_tool_state_root_entry "${entry_name}"; then
        printf '%s\n' '--purge-tool-state'
        return 0
    fi
    if repo_hygiene_is_scratch_root_entry "${entry_name}"; then
        printf '%s\n' '--purge-tmp'
        return 0
    fi
    printf '%s\n' '(none)'
}

repo_hygiene_list_root_entries() {
    local repo_root=$1
    local path
    local -a entries=()

    shopt -s nullglob
    entries=("${repo_root}"/* "${repo_root}"/.[!.]* "${repo_root}"/..?*)
    shopt -u nullglob

    for path in "${entries[@]}"; do
        printf '%s\n' "${path##*/}"
    done | LC_ALL=C sort -u
}

repo_hygiene_print_local_state_report() {
    local repo_root=$1
    local local_state_root
    local size
    local category
    local cleanup_flag

    for local_state_root in "${repo_hygiene_local_state_root_entries[@]}"; do
        [[ -e "${repo_root}/${local_state_root}" ]] || continue
        size="$(du -sh "${repo_root}/${local_state_root}" | awk '{print $1}')"
        category="$(repo_hygiene_local_state_category "${local_state_root}")"
        cleanup_flag="$(repo_hygiene_local_state_cleanup_flag "${local_state_root}")"
        printf '%s\t%s\t%s\t%s\n' \
            "${size}" \
            "${category}" \
            "${cleanup_flag}" \
            "${repo_root}/${local_state_root}"
    done |
        sort -h -k1,1 |
        while IFS=$'\t' read -r size category cleanup_flag path; do
            printf '  %-6s %-16s %-24s %s\n' "${size}" "${category}" "${cleanup_flag}" "${path}"
        done
}
