#!/usr/bin/env bash
# Keep the contract replacement split enforced: no resurrected protocol module,
# canonical module graph intact, and the accepted ADR surfaced in developer docs.

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
readonly settings_file="${repo_root}/settings.gradle.kts"
readonly authoring_build_file="${repo_root}/authoring-java/build.gradle.kts"
readonly cli_build_file="${repo_root}/cli/build.gradle.kts"
readonly cli_module_file="${repo_root}/cli/src/main/java/module-info.java"
readonly authoring_module_file="${repo_root}/authoring-java/src/main/java/module-info.java"
readonly contract_module_file="${repo_root}/contract/src/main/java/module-info.java"
readonly excel_foundation_module_file="${repo_root}/excel-foundation/src/main/java/module-info.java"
readonly executor_module_file="${repo_root}/executor/src/main/java/module-info.java"
readonly engine_module_file="${repo_root}/engine/src/main/java/module-info.java"
readonly jazzer_conventions_file="${repo_root}/gradle/build-logic/src/main/kotlin/dev/erst/gridgrind/buildlogic/GridGrindJazzerConventionsPlugin.kt"
readonly developer_doc="${repo_root}/docs/DEVELOPER.md"
readonly adr_doc="${repo_root}/docs/DEVELOPER_CONTRACT_REPLACEMENT_ADR.md"

[[ -f "${settings_file}" ]] || die "missing settings.gradle.kts"
[[ -d "${repo_root}/authoring-java" ]] || die "missing authoring-java module directory"
[[ -d "${repo_root}/contract" ]] || die "missing contract module directory"
[[ -d "${repo_root}/excel-foundation" ]] || die "missing excel-foundation module directory"
[[ -d "${repo_root}/executor" ]] || die "missing executor module directory"
[[ ! -d "${repo_root}/protocol" ]] || die "legacy top-level protocol module still exists"

grep -Fq 'include("excel-foundation", "engine", "contract", "executor", "authoring-java", "cli")' "${settings_file}" || die \
    "settings.gradle.kts does not include the canonical six-module product graph"
grep -Fq '"protocol"' "${settings_file}" && die \
    "settings.gradle.kts still includes the deleted protocol module"

grep -Fq 'api(project(":contract"))' "${authoring_build_file}" || die \
    "authoring-java no longer depends on contract as its canonical authoring boundary"
grep -Fq 'api(project(":executor"))' "${authoring_build_file}" && die \
    "authoring-java still depends on executor despite the contract-only authoring boundary"
grep -Fq 'implementation(project(":executor"))' "${cli_build_file}" || die \
    "cli no longer depends on executor"
grep -Fq 'implementation(project(":protocol"))' "${cli_build_file}" && die \
    "cli still depends on the deleted protocol module"

grep -Fq 'module dev.erst.gridgrind.authoring {' "${authoring_module_file}" || die \
    "authoring-java module-info does not declare the canonical authoring module name"
grep -Fq 'requires transitive dev.erst.gridgrind.contract;' "${authoring_module_file}" || die \
    "authoring-java module-info no longer depends transitively on contract"
grep -Fq 'requires transitive dev.erst.gridgrind.executor;' "${authoring_module_file}" && die \
    "authoring-java module-info still depends transitively on executor"
grep -Fq 'requires dev.erst.gridgrind.executor;' "${cli_module_file}" || die \
    "cli module-info no longer requires executor"
grep -Fq 'requires dev.erst.gridgrind.protocol;' "${cli_module_file}" && die \
    "cli module-info still requires the deleted protocol module"

grep -Fq 'module dev.erst.gridgrind.contract {' "${contract_module_file}" || die \
    "contract module-info does not declare the canonical contract module name"
grep -Fq 'module dev.erst.gridgrind.excel.foundation {' "${excel_foundation_module_file}" || die \
    "excel-foundation module-info does not declare the canonical foundation module name"
grep -Fq 'module dev.erst.gridgrind.executor {' "${executor_module_file}" || die \
    "executor module-info does not declare the executor module name"
grep -Fq 'requires transitive dev.erst.gridgrind.contract;' "${executor_module_file}" || die \
    "executor module-info no longer depends transitively on contract"
grep -Fq 'requires transitive dev.erst.gridgrind.excel.foundation;' "${contract_module_file}" || die \
    "contract module-info no longer exposes the shared excel-foundation module transitively"
grep -Fq 'requires dev.erst.gridgrind.excel.foundation;' "${engine_module_file}" || die \
    "engine module-info no longer depends on the shared excel-foundation module"
grep -Fq '"dev.erst.gridgrind:executor"' "${jazzer_conventions_file}" || die \
    "jazzer build logic no longer consumes executor as the contract execution bridge"
grep -Fq '"dev.erst.gridgrind:protocol"' "${jazzer_conventions_file}" && die \
    "jazzer build logic still depends on the deleted protocol module"

grep -Fq 'DEVELOPER_CONTRACT_REPLACEMENT_ADR.md' "${developer_doc}" || die \
    "developer reference no longer links the accepted contract-replacement ADR"
grep -Fq 'dev.erst.gridgrind.authoring -> dev.erst.gridgrind.contract -> dev.erst.gridgrind.excel.foundation' "${developer_doc}" || die \
    "developer reference no longer documents the authoring-java module graph"
grep -Fq 'dev.erst.gridgrind.cli -> dev.erst.gridgrind.executor -> dev.erst.gridgrind.contract -> dev.erst.gridgrind.excel.foundation' "${developer_doc}" || die \
    "developer reference no longer documents the canonical module graph"
grep -Fq 'dev.erst.gridgrind.executor -> dev.erst.gridgrind.engine -> dev.erst.gridgrind.excel.foundation' "${developer_doc}" || die \
    "developer reference no longer documents the engine-to-foundation bridge graph"
grep -Fq '**Status**: Accepted' "${adr_doc}" || die \
    "contract replacement ADR is missing its accepted status"

if rg -n \
    'dev\.erst\.gridgrind\.protocol|project\(":protocol"\)|module dev\.erst\.gridgrind\.protocol|requires dev\.erst\.gridgrind\.protocol|protocol/src/|protocol/build\.gradle\.kts' \
    "${repo_root}/README.md" \
    "${repo_root}/docs" \
    "${repo_root}/scripts" \
    "${repo_root}/gradle" \
    "${repo_root}/settings.gradle.kts" \
    "${repo_root}/check.sh" \
    "${repo_root}/cli" \
    "${repo_root}/contract" \
    "${repo_root}/executor" \
    "${repo_root}/engine" \
    "${repo_root}/.github" \
    -g'!**/build/**' -g'!**/.gradle/**' -g'!scripts/test-contract-module-split.sh' >/dev/null; then
    die "legacy protocol module references still exist in the live product/build/doc surface"
fi

printf 'contract-module-split regression: success\n'
