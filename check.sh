#!/usr/bin/env bash
# Run all local verification gates and release packaging checks.
#
# This file intentionally lives at the repository root beside gradlew because it is the canonical
# human-facing project gate entrypoint. The scripts/ directory is reserved for subordinate helper
# scripts that workflows and this root gate invoke.
#
# Stage 1 runs all root-project quality gates and generates coverage reports for local inspection:
#   check    -> Spotless (format), Error Prone (compile-time checks), PMD (static analysis),
#               JaCoCo coverage thresholds, and all unit tests
#   coverage -> per-module and aggregated JaCoCo HTML/XML reports
#
# Stage 2 runs the nested Jazzer verification build:
#   jazzer check -> deterministic Jazzer support tests plus committed-seed regression replay
#
# CI runs only the root check task (verification without report generation).
#
# Stage 3 mirrors the GitHub release packaging workflow:
#   :cli:shadowJar -> build the distributable fat JAR
#
# Stage 4 syntax-checks the release-surface shell scripts:
#   bash -n check.sh scripts/*.sh
#
# Stage 5 exercises the Docker release surface from a non-default working directory:
#   scripts/docker-smoke.sh -> build the image and verify help/request/response/save behavior
#
# The script is location-independent: it always targets the repository that contains this file,
# even when invoked from another working directory or through a symlink.
#
# Local runs keep the Gradle daemon for speed. When CI is set, the script adds --no-daemon
# automatically to match the GitHub workflows. Non-interactive runs use --console=plain unless
# the caller already selected a console mode.
#
# Exit status: 0 on success. Any failing Gradle stage or script precondition returns a non-zero
# exit status. The script emits a final human-readable result line plus one machine-readable
# summary line: [CHECK-SUMMARY] status=<success|failure> stage=<stage-id> exit_code=<n>
#
# Usage: ./check.sh [supported gradle options]

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

readonly repo_root="$(resolve_script_dir)"
readonly gradlew="${repo_root}/gradlew"
current_stage_id='startup'
current_stage_label='starting'
emit_final_status_enabled=true

print_usage() {
    printf '%s\n' \
        'Usage: ./check.sh [supported gradle options]' \
        '' \
        'Runs five fixed stages against the repository that contains this script:' \
        '  1. check coverage' \
        '  2. jazzer check' \
        '  3. :cli:shadowJar' \
        '  4. bash -n check.sh scripts/*.sh' \
        '  5. scripts/docker-smoke.sh' \
        '' \
        'Supported options:' \
        '  -h, --help' \
        '  --console=plain|auto|rich|verbose' \
        '  --console plain|auto|rich|verbose' \
        '  --warning-mode=all|fail|summary|none' \
        '  --warning-mode all|fail|summary|none' \
        '  --daemon, --no-daemon' \
        '  --dry-run, -m' \
        '  --stacktrace, --full-stacktrace, -s, -S' \
        '  --info, --debug, --warn, --quiet, -i, -q' \
        '  --scan, --profile, --continue, --no-continue' \
        '  --parallel, --no-parallel' \
        '  --build-cache, --no-build-cache' \
        '  --configuration-cache, --no-configuration-cache' \
        '  --rerun-tasks, --refresh-dependencies, --offline' \
        '  -Dname=value, -Pname=value' \
        '' \
        'Unsupported inputs:' \
        '  - positional Gradle tasks/selectors such as help, test, tasks, or :cli:test' \
        '  - project-location overrides such as --project-dir, --build-file, or --settings-file' \
        '' \
        'Diagnostic escalation:' \
        '  - Use ./check.sh --info ONLY for normal project-verification failures when the default output' \
        '    does not yet show enough assertion detail or task context to fix the code.' \
        '  - Use ./check.sh --stacktrace ONLY for build-tool, plugin, environment, or filesystem' \
        '    failures when the default output does not already point to the failing source location.' \
        '' \
        'For anything outside this fixed interface, run ./gradlew directly.'
}

print_failure_guidance() {
    case "${current_stage_id}" in
        argument-validation)
            printf '%s\n' \
                'Diagnostic escalation:' \
                '  - Do not retry with --info or --stacktrace.' \
                '  - Fix the script arguments first; this failure happened before Gradle started.'
            ;;
        *)
            printf '%s\n' \
                'Diagnostic escalation:' \
                '  - Retry with ./check.sh --info ONLY if this is a normal verification failure and the' \
                '    default output still lacks the exact assertion message, expected/actual values, or' \
                '    enough task context to fix the code.' \
                '  - Use ./check.sh --stacktrace ONLY if this looks like build infrastructure rather than' \
                '    a normal code defect: AccessDeniedException, file-lock or configuration-cache errors,' \
                '    plugin or worker crashes, toolchain failures, or task exceptions without an' \
                '    actionable source file and line.' \
                '  - Do not enable either flag by default; use them only as targeted escalation.'
            ;;
    esac
}

emit_final_status() {
    local exit_code=$?
    [[ "${emit_final_status_enabled}" == true ]] || return 0
    if [[ "${exit_code}" -eq 0 ]]; then
        printf 'Result: success\n'
        printf '[CHECK-SUMMARY] status=success stage=%s exit_code=%d\n' "${current_stage_id}" "${exit_code}"
    else
        printf 'Result: failure during %s\n' "${current_stage_label}"
        print_failure_guidance
        printf '[CHECK-SUMMARY] status=failure stage=%s exit_code=%d\n' "${current_stage_id}" "${exit_code}"
    fi
}

trap emit_final_status EXIT

[[ -x "${gradlew}" ]] || die "missing executable Gradle wrapper at ${gradlew}"
[[ -f "${repo_root}/settings.gradle.kts" ]] || die "missing Gradle settings file at ${repo_root}/settings.gradle.kts"

current_stage_id='argument-validation'
current_stage_label='argument validation'

gradle_args=()
has_daemon_flag=false
has_console_flag=false
expects_value=''
for gradle_arg in "$@"; do
    if [[ -n "${expects_value}" ]]; then
        if [[ "${gradle_arg}" == -* ]]; then
            die "option ${expects_value} requires a value"
        fi
        gradle_args+=("${gradle_arg}")
        expects_value=''
        continue
    fi
    case "${gradle_arg}" in
        -h|--help)
            emit_final_status_enabled=false
            print_usage
            exit 0
            ;;
        --daemon|--no-daemon)
            has_daemon_flag=true
            gradle_args+=("${gradle_arg}")
            ;;
        --console)
            has_console_flag=true
            gradle_args+=("${gradle_arg}")
            expects_value='--console'
            ;;
        --console=*)
            has_console_flag=true
            gradle_args+=("${gradle_arg}")
            ;;
        --warning-mode)
            gradle_args+=("${gradle_arg}")
            expects_value='--warning-mode'
            ;;
        --warning-mode=*|--dry-run|--stacktrace|--full-stacktrace|--info|--debug|--warn|--quiet|--scan|--profile|--continue|--no-continue|--parallel|--no-parallel|--build-cache|--no-build-cache|--configuration-cache|--no-configuration-cache|--rerun-tasks|--refresh-dependencies|--offline|-m|-q|-i|-s|-S|-D*|-P*)
            gradle_args+=("${gradle_arg}")
            ;;
        -p|--project-dir|--project-dir=*|-b|--build-file|--build-file=*|-c|--settings-file|--settings-file=*)
            die "do not override the project location; this script always targets ${repo_root}"
            ;;
        -x|--exclude-task|--exclude-task=*|--include-build|--include-build=*|--tests|--tests=*)
            die "unsupported Gradle option for this fixed script interface: ${gradle_arg}"
            ;;
        -*)
            die "unsupported Gradle option for this fixed script interface: ${gradle_arg}"
            ;;
        *)
            die "positional Gradle tasks or selectors are not supported here: ${gradle_arg}"
            ;;
    esac
done

[[ -z "${expects_value}" ]] || die "option ${expects_value} requires a value"

if [[ -n "${CI:-}" && "${has_daemon_flag}" == false ]]; then
    gradle_args+=(--no-daemon)
fi

if [[ ( -n "${CI:-}" || ! -t 1 ) && "${has_console_flag}" == false ]]; then
    gradle_args+=(--console=plain)
fi

run_stage() {
    local stage_id=$1
    local stage_label=$2
    local project_dir=$3
    shift 3
    current_stage_id="${stage_id}"
    current_stage_label="${stage_label}"
    printf '%s\n' "${stage_label}"
    "${gradlew}" --project-dir "${project_dir}" "$@" "${gradle_args[@]}"
}

run_shell_stage() {
    local stage_id=$1
    local stage_label=$2
    shift 2
    current_stage_id="${stage_id}"
    current_stage_label="${stage_label}"
    printf '%s\n' "${stage_label}"
    (cd "${repo_root}" && "$@")
}

run_stage 'quality-gates' 'Stage 1/5: running quality gates' "${repo_root}" check coverage
run_stage \
    'jazzer-check' \
    'Stage 2/5: running Jazzer support tests and regression replay' \
    "${repo_root}/jazzer" \
    check
run_stage 'cli-shadowjar' 'Stage 3/5: building CLI fat JAR' "${repo_root}" :cli:shadowJar

shell_syntax_targets=("${repo_root}/check.sh")
if [[ -d "${repo_root}/scripts" ]]; then
    while IFS= read -r shell_script_path; do
        shell_syntax_targets+=("${shell_script_path}")
    done < <(find "${repo_root}/scripts" -maxdepth 1 -type f -name '*.sh' | sort)
fi

run_shell_stage 'shell-syntax' 'Stage 4/5: syntax-checking release-surface shell scripts' \
    bash -n "${shell_syntax_targets[@]}"
run_shell_stage 'docker-smoke' 'Stage 5/5: running Docker smoke test' "${repo_root}/scripts/docker-smoke.sh"
