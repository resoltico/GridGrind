#!/usr/bin/env bash
# Guard the dedicated mutation-check entrypoint and Gradle ownership boundary.

set -euo pipefail

die() { printf 'error: %s\n' "$1" >&2; exit 1; }

readonly script_dir="$(cd -P -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
readonly repo_root="$(cd -P -- "${script_dir}/.." && pwd)"
readonly mutation_check="${repo_root}/check_mutation.sh"
readonly mutation_plugin="${repo_root}/gradle/build-logic/src/main/kotlin/dev/erst/gridgrind/buildlogic/GridGrindMutationConventionsPlugin.kt"
readonly mutation_workflow="${repo_root}/.github/workflows/mutation.yml"
readonly root_conventions="${repo_root}/gradle/build-logic/src/main/kotlin/dev/erst/gridgrind/buildlogic/GridGrindRootConventionsPlugin.kt"
readonly mutation_verification="${repo_root}/gradle/build-logic/src/main/kotlin/dev/erst/gridgrind/buildlogic/GridGrindMutationVerification.kt"
readonly contract_build="${repo_root}/contract/build.gradle.kts"
readonly engine_build="${repo_root}/engine/build.gradle.kts"
readonly cli_build="${repo_root}/cli/build.gradle.kts"
readonly executor_build="${repo_root}/executor/build.gradle.kts"

[[ -x "${mutation_check}" ]] || die "check_mutation.sh must be executable"
bash -n "${mutation_check}"
"${mutation_check}" --help | grep -Fq 'mutationCheck' || die "mutation help no longer names the fixed Gradle task"
grep -Fq 'tasks.register("mutationCheck")' "${mutation_verification}" || die "root mutation verification no longer owns the aggregate task"
grep -Fq 'PitestPluginExtension' "${mutation_plugin}" || die "mutation convention no longer owns PIT configuration"
grep -Fq 'cleanPitestReport' "${mutation_plugin}" || die "mutation convention no longer cleans stale reports"
grep -Fq 'verifyPitestScope' "${mutation_plugin}" || die "mutation convention no longer verifies every configured scope pattern"
grep -Fq 'verifyPitestReport' "${mutation_plugin}" || die "mutation convention no longer rejects non-killed PIT outcomes"
grep -Fq 'mutators.set(setOf("STRONGER"))' "${mutation_plugin}" || die "mutation convention no longer pins the reviewed stronger mutator group"
grep -Fq 'timeoutConstInMillis.set(10_000)' "${mutation_plugin}" || die "mutation convention no longer pins a generous PIT timeout constant"
grep -Fq 'timeoutFactor.set(BigDecimal("3.0"))' "${mutation_plugin}" || die "mutation convention no longer pins a PIT timeout factor"
if grep -Fq 'rootProject.tasks.maybeCreate("mutationCheck")' "${mutation_plugin}"; then
    die "mutation convention must not create the root aggregate from a subproject"
fi
grep -Fq 'registerGridGrindMutationCheck()' "${root_conventions}" || die "root conventions no longer own mutation aggregate registration"
grep -Fq '":contract:verifyPitestReport"' "${mutation_verification}" || die "mutation aggregate no longer requires contract verification"
grep -Fq '":engine:verifyPitestReport"' "${mutation_verification}" || die "mutation aggregate no longer requires engine verification"
grep -Fq '":cli:verifyPitestReport"' "${mutation_verification}" || die "mutation aggregate no longer requires CLI verification"
grep -Fq '":executor:verifyPitestReport"' "${mutation_verification}" || die "mutation aggregate no longer requires executor verification"
grep -Fq 'id("gridgrind.mutation-conventions")' "${contract_build}" || die "contract no longer applies the mutation convention"
grep -Fq 'id("gridgrind.mutation-conventions")' "${engine_build}" || die "engine no longer applies the mutation convention"
grep -Fq 'id("gridgrind.mutation-conventions")' "${cli_build}" || die "CLI no longer applies the mutation convention"
grep -Fq 'id("gridgrind.mutation-conventions")' "${executor_build}" || die "executor no longer applies the mutation convention"
grep -Fq -- '--no-parallel' "${mutation_check}" || die "mutation wrapper no longer serializes module PIT tasks"
grep -Fq 'mustRunAfter(":contract:pitest")' "${engine_build}" || die "engine PIT no longer follows contract PIT"
grep -Fq 'mustRunAfter(":engine:pitest")' "${executor_build}" || die "architecture PIT no longer follows engine PIT"
if "${mutation_check}" clean >/dev/null 2>&1; then
    die "mutation wrapper accepted an additional positional Gradle task"
fi
if "${mutation_check}" --project-dir=/tmp >/dev/null 2>&1; then
    die "mutation wrapper accepted a project-location override"
fi
grep -Fq 'pull_request:' "${mutation_workflow}" || die "mutation workflow no longer validates relevant PRs"
grep -Fq 'schedule:' "${mutation_workflow}" || die "mutation workflow no longer runs on schedule"
grep -Fq 'workflow_dispatch:' "${mutation_workflow}" || die "mutation workflow no longer supports manual runs"
grep -Fq 'merge_group:' "${mutation_workflow}" || die "mutation workflow no longer validates merge-queue batches"
grep -Fq 'if: always()' "${mutation_workflow}" || die "mutation reports are no longer retained on failure"
grep -Fq 'actions/upload-artifact@043fb46d1a93c77aae656e7c1c64a875d1fc6a0a' "${mutation_workflow}" || die "mutation report upload action is not pinned"
grep -Fq 'cache-read-only: ${{ github.event_name == '\''pull_request'\'' }}' "${mutation_workflow}" || die "mutation PRs can poison the Gradle cache"
grep -Fq 'contract/build/reports/pitest' "${mutation_workflow}" || die "contract mutation report is not retained"
grep -Fq 'engine/build/reports/pitest' "${mutation_workflow}" || die "engine mutation report is not retained"
grep -Fq 'cli/build/reports/pitest' "${mutation_workflow}" || die "CLI mutation report is not retained"
grep -Fq 'executor/build/reports/pitest' "${mutation_workflow}" || die "architecture mutation report is not retained"
grep -Fq "'contract/src/**'" "${mutation_workflow}" || die "mutation workflow no longer safely covers contract scope changes"
grep -Fq "'engine/src/**'" "${mutation_workflow}" || die "mutation workflow no longer safely covers engine scope changes"
grep -Fq "'cli/src/**'" "${mutation_workflow}" || die "mutation workflow no longer safely covers CLI scope changes"
grep -Fq "'executor/src/**'" "${mutation_workflow}" || die "mutation workflow no longer safely covers executor scope changes"
grep -Fq 'if-no-files-found: error' "${mutation_workflow}" || die "mutation workflow must fail when a required report is absent"
grep -Fq 'retention-days: 30' "${mutation_workflow}" || die "mutation report retention drifted"
printf 'mutation check regression: success\n'
