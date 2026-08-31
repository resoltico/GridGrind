#!/usr/bin/env bash
# Canonical fixed-stage contract for the root check.sh entrypoint.

readonly check_stage_ids=(
    repo-hygiene
    quality-gates
    jazzer-check
    cli-shadowjar
    shell-syntax
    docker-smoke
)

readonly check_stage_labels=(
    'Stage 1/6: verifying repo hygiene'
    'Stage 2/6: running quality gates'
    'Stage 3/6: running Jazzer support tests and regression replay'
    'Stage 4/6: building CLI fat JAR'
    'Stage 5/6: checking release-surface shell scripts'
    'Stage 6/6: running Docker smoke test'
)

readonly check_stage5_script_paths=(
    scripts/run-quality-gates.sh
    scripts/repo-hygiene-support.sh
    scripts/verify-repo-hygiene.sh
    scripts/clean-repo-hygiene.sh
    scripts/test-check-process-support.sh
    scripts/test-check-file-support.sh
    scripts/test-check-stage-contract.sh
    scripts/test-check-mutation.sh
    scripts/test-cli-distribution-surface.sh
    scripts/test-cli-shadow-jar-support.sh
    scripts/test-contract-module-split.sh
    scripts/test-documentation-contract.sh
    scripts/test-explicit-import-gate.sh
    scripts/test-jazzer-public-surface.sh
    scripts/test-jazzer-run-lock.sh
    scripts/test-jazzer-portable-stat.sh
    scripts/test-jazzer-tool-task-cache.sh
    scripts/test-jazzer-wrapper-ux.sh
    scripts/test-jazzer-wrapper-arg-order.sh
    scripts/test-repo-verification-lock.sh
    scripts/test-selector-contract-surface.sh
    scripts/test-verify-release-merge-handoff.sh
    scripts/test-verify-release-candidate-tag.sh
    scripts/test-verify-release-primary-checkout.sh
    scripts/test-verify-cli-contract.sh
    scripts/test-verify-cli-discovery-execution.sh
    scripts/verify-cli-discovery-execution.sh
    scripts/test-verify-container-publication.sh
    scripts/test-publication-contract.sh
)

check_stage_usage_lines() {
    printf '%s\n' \
        '  1. scripts/verify-repo-hygiene.sh' \
        '  2. check coverage' \
        '  3. jazzer check' \
        '  4. :cli:shadowJar' \
        "  5. $(check_stage5_usage_command)" \
        '  6. scripts/docker-smoke.sh'
}

check_stage5_usage_command() {
    local command='bash -n check.sh check_mutation.sh scripts/*.sh jazzer/bin/* && scripts/verify-cli-contract.sh jar ./cli/build/libs/gridgrind.jar'
    local script_path=''
    for script_path in "${check_stage5_script_paths[@]}"; do
        command="${command} && ${script_path}"
    done
    printf '%s\n' "${command}"
}

check_stage_execute() {
    local stage_id=$1
    local stage_label=$2
    local check_repo_root=$3

    case "${stage_id}" in
        repo-hygiene)
            run_shell_stage "${stage_id}" "${stage_label}" \
                "${check_repo_root}/scripts/verify-repo-hygiene.sh" \
                --repo-root "${check_repo_root}"
            ;;
        quality-gates)
            run_stage "${stage_id}" "${stage_label}" "${check_repo_root}" check coverage
            ;;
        jazzer-check)
            run_stage "${stage_id}" "${stage_label}" "${check_repo_root}/jazzer" check
            ;;
        cli-shadowjar)
            run_stage "${stage_id}" "${stage_label}" "${check_repo_root}" :cli:shadowJar
            ;;
        shell-syntax)
            local shell_syntax_targets=("${check_repo_root}/check.sh" "${check_repo_root}/check_mutation.sh")
            local shell_script_path=''
            if [[ -d "${check_repo_root}/scripts" ]]; then
                while IFS= read -r shell_script_path; do
                    shell_syntax_targets+=("${shell_script_path}")
                done < <(find "${check_repo_root}/scripts" -maxdepth 1 -type f -name '*.sh' | sort)
            fi
            if [[ -d "${check_repo_root}/jazzer/bin" ]]; then
                while IFS= read -r shell_script_path; do
                    shell_syntax_targets+=("${shell_script_path}")
                done < <(find "${check_repo_root}/jazzer/bin" -maxdepth 1 -type f | sort)
            fi
            run_shell_stage "${stage_id}" "${stage_label}" \
                bash -c '
                    set -euo pipefail
                    # shellcheck source=/dev/null
                    source "'"${check_repo_root}"'/scripts/lib/cli-shadow-jar-support.sh"
                    shell_syntax_targets=()
                    while [[ "$1" != "--" ]]; do
                        shell_syntax_targets+=("$1")
                        shift
                    done
                    shift
                    bash -n "${shell_syntax_targets[@]}"
                    cli_jar_path="$(ensure_cli_shadow_jar "'"${check_repo_root}"'")"
                    "'"${check_repo_root}"'/scripts/verify-cli-contract.sh" jar "${cli_jar_path}" >/dev/null
                    local_script_path=""
                    for local_script_path in "$@"; do
                        bash "'"${check_repo_root}"'/${local_script_path}"
                    done
                ' bash "${shell_syntax_targets[@]}" -- "${check_stage5_script_paths[@]}"
            ;;
        docker-smoke)
            run_shell_stage "${stage_id}" "${stage_label}" "${check_repo_root}/scripts/docker-smoke.sh"
            ;;
        *)
            die "unsupported fixed stage id from ${stage_contract_script}: ${stage_id}"
            ;;
    esac
}
