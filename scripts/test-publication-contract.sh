#!/usr/bin/env bash
# Keep the release-publication surface deterministic: pinned base image, guarded workflow-dispatch
# publishing, explicit attestations, accurate OCI labels, dynamic root coverage wiring, and a
# narrow contract compile API.

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
readonly gitignore_file="${repo_root}/.gitignore"
readonly gitattributes_file="${repo_root}/.gitattributes"
readonly dockerignore_file="${repo_root}/.dockerignore"
readonly dockerfile="${repo_root}/Dockerfile"
readonly release_workflow="${repo_root}/.github/workflows/release.yml"
readonly container_workflow="${repo_root}/.github/workflows/container.yml"
readonly root_plugin="${repo_root}/gradle/build-logic/src/main/kotlin/dev/erst/gridgrind/buildlogic/GridGrindRootConventionsPlugin.kt"
readonly contract_build="${repo_root}/contract/build.gradle.kts"
readonly cli_jar="${repo_root}/cli/build/libs/gridgrind.jar"
readonly docker_smoke_script="${repo_root}/scripts/docker-smoke.sh"
readonly container_verify_script="${repo_root}/scripts/verify-container-publication.sh"
readonly stage_contract_script="${repo_root}/scripts/check-stage-contract.sh"
readonly release_protocol_doc="${repo_root}/docs/RELEASE_PROTOCOL.md"
readonly temp_parent="${repo_root}/tmp/test-publication-contract"
test_root=''

cleanup() {
    [[ -n "${test_root}" ]] && rm -rf "${test_root}" || true
}

trap cleanup EXIT

fixed_pattern_exists() {
    local pattern=$1
    local path=$2

    if command -v rg >/dev/null 2>&1; then
        rg -Fq -- "${pattern}" "${path}"
        return $?
    fi
    grep -Fq -- "${pattern}" "${path}"
}

grep -Eq \
    '^FROM azul/zulu-openjdk-alpine:26-jre@sha256:[0-9a-f]{64}$' \
    "${dockerfile}" || die "Dockerfile base image is not digest-pinned"

git -C "${repo_root}" check-ignore -q AGENTS.md && die \
    "root AGENTS.md is still ignored, so agent instructions cannot be tracked"
git -C "${repo_root}" check-ignore -q .codex/AGENTS_EXTRA.md && die \
    "repo-owned /.codex/ content is still ignored, so agent instructions cannot be tracked"

grep -Fq '/AGENTS.md   export-ignore' "${gitattributes_file}" || die \
    ".gitattributes no longer excludes /AGENTS.md from source archives"
grep -Fq '/.codex      export-ignore' "${gitattributes_file}" || die \
    ".gitattributes no longer excludes /.codex from source archives"
grep -Fq '/.codex/**   export-ignore' "${gitattributes_file}" || die \
    ".gitattributes no longer excludes /.codex/** from source archives"

grep -Eq '^!.*AGENTS\.md$' "${dockerignore_file}" && die \
    ".dockerignore unexpectedly whitelists /AGENTS.md into the Docker build context"
grep -Eq '^!.*\.codex(/|\*\*|$)' "${dockerignore_file}" && die \
    ".dockerignore unexpectedly whitelists /.codex into the Docker build context"

command -v jar >/dev/null 2>&1 || die "jar is required for publication contract verification"
[[ -f "${cli_jar}" ]] || die "missing CLI fat JAR at ${cli_jar}"
jar_listing="$(jar tf "${cli_jar}")"
grep -Fq 'AGENTS.md' <<<"${jar_listing}" && die \
    "CLI fat JAR unexpectedly contains /AGENTS.md"
grep -Fq '.codex/' <<<"${jar_listing}" && die \
    "CLI fat JAR unexpectedly contains /.codex/"

mkdir -p "${temp_parent}"
test_root="${temp_parent}/run.$$"
rm -rf "${test_root}"
mkdir -p "${test_root}/archive-root/.codex/protocol" "${test_root}/archive-root/src"
cp "${gitattributes_file}" "${test_root}/archive-root/.gitattributes"
printf '# synthetic agent entry point\n' > "${test_root}/archive-root/AGENTS.md"
printf '# synthetic codex doc\n' > "${test_root}/archive-root/.codex/AGENTS_EXTRA.md"
printf '# synthetic nested codex doc\n' > "${test_root}/archive-root/.codex/protocol/guide.md"
printf 'public file\n' > "${test_root}/archive-root/src/published.txt"
git -C "${test_root}/archive-root" init >/dev/null
git -C "${test_root}/archive-root" config user.name "GridGrind Test"
git -C "${test_root}/archive-root" config user.email "gridgrind-test@example.com"
git -C "${test_root}/archive-root" add .gitattributes AGENTS.md .codex/AGENTS_EXTRA.md .codex/protocol/guide.md src/published.txt
git -C "${test_root}/archive-root" commit -m "Archive surface fixture" >/dev/null
git -C "${test_root}/archive-root" archive --format=tar --output "${test_root}/archive.tar" HEAD
archive_listing="$(tar -tf "${test_root}/archive.tar")"
grep -Fq 'src/published.txt' <<<"${archive_listing}" || die \
    "git archive no longer includes ordinary tracked files for the public source asset"
grep -Fq 'AGENTS.md' <<<"${archive_listing}" && die \
    "git archive still includes /AGENTS.md in the public source asset"
grep -Fq '.codex/AGENTS_EXTRA.md' <<<"${archive_listing}" && die \
    "git archive still includes /.codex/AGENTS_EXTRA.md in the public source asset"
grep -Fq '.codex/protocol/guide.md' <<<"${archive_listing}" && die \
    "git archive still includes nested /.codex content in the public source asset"

grep -Fq './scripts/verify-release-candidate-tag.sh "${{ steps.target-tag.outputs.tag }}"' \
    "${release_workflow}" || die "release workflow does not enforce the shared tag verifier"
grep -Fq './scripts/verify-release-candidate-tag.sh "${{ steps.target-tag.outputs.tag }}"' \
    "${container_workflow}" || die "container workflow does not enforce the shared tag verifier"
grep -Fq './scripts/verify-cli-contract.sh jar ./cli/build/libs/gridgrind.jar' \
    "${release_workflow}" || die "release workflow does not verify the packaged CLI contract"
grep -Fq 'gradle_command=(' "${docker_smoke_script}" || die \
    "docker smoke no longer prepares a dedicated Gradle invocation for the packaged CLI fat JAR"
grep -Fq '"${repo_root}/gradlew"' "${docker_smoke_script}" || die \
    "docker smoke no longer invokes the repository Gradle wrapper"
grep -Fq -- '--console=plain' "${docker_smoke_script}" || die \
    "docker smoke no longer forces plain console output for the packaged CLI rebuild"
grep -Fq -- '--no-daemon' "${docker_smoke_script}" || die \
    "docker smoke no longer disables the Gradle daemon for the packaged CLI rebuild"
grep -Fq 'gradle_command+=(:cli:shadowJar)' "${docker_smoke_script}" || die \
    "docker smoke no longer rebuilds the packaged CLI fat JAR"
grep -Fq '"${repo_root}/scripts/verify-cli-contract.sh" docker-image "${image_tag}"' \
    "${docker_smoke_script}" || die "docker smoke no longer verifies the local image CLI contract"
grep -Fq "readonly streaming_read_request_rel='requests odd/request streaming readback [docker #smoke].json'" \
    "${docker_smoke_script}" || die "docker smoke no longer stages a separate streaming readback request"
grep -Fq -- '--request "${streaming_read_request_rel}"' "${docker_smoke_script}" || die \
    "docker smoke no longer reopens the streamed workbook for the summary readback"
grep -Fq '"level": "VERBOSE"' "${docker_smoke_script}" || die \
    "docker smoke no longer exercises verbose execution journaling from the packaged artifact"
grep -Fq 'docker smoke response did not include the structured execution journal' "${docker_smoke_script}" || die \
    "docker smoke no longer asserts response-journal presence"
grep -Fq 'docker smoke create request did not stream live verbose journal events to stderr' "${docker_smoke_script}" || die \
    "docker smoke no longer asserts live verbose stderr journal streaming"
grep -Fq 'run_verify_cli_contract "${image_name}:${expected_version}"' \
    "${container_verify_script}" || die "public container verification no longer checks the version tag contract"
grep -Fq 'run_verify_cli_contract "${image_name}:latest"' \
    "${container_verify_script}" || die "public container verification no longer checks the latest tag contract"
grep -Fq 'scripts/test-verify-release-primary-checkout.sh' "${stage_contract_script}" || die \
    "Stage 4 contract no longer exercises the release primary-checkout regression"
grep -Fq './scripts/verify-release-primary-checkout.sh "$PRIMARY_CHECKOUT" "X.Y.Z"' \
    "${release_protocol_doc}" || die "release protocol no longer requires the primary-checkout closeout verifier"

grep -Fq 'checks: read' "${release_workflow}" || die "release workflow is missing checks: read permission"
grep -Fq 'checks: read' "${container_workflow}" || die "container workflow is missing checks: read permission"

grep -Fq 'provenance: mode=max' "${container_workflow}" || die "container workflow does not publish explicit provenance"
grep -Fq 'sbom: true' "${container_workflow}" || die "container workflow does not publish an SBOM attestation"
grep -Fqx '            org.opencontainers.image.licenses=MIT AND Apache-2.0 AND BSD-3-Clause' \
    "${container_workflow}" || die "container workflow still uses the wrong OCI license label"
grep -Fqx '            org.opencontainers.image.licenses=MIT' "${container_workflow}" && die \
    "container workflow still contains the legacy MIT-only OCI license label"

grep -Fq 'api(libs.jackson.annotations)' "${contract_build}" || die \
    "contract no longer exposes the Jackson annotations API needed by downstream compiles"
grep -Fq 'implementation(libs.jackson.databind)' "${contract_build}" || die \
    "contract no longer keeps jackson-databind off consumer compile classpaths"
grep -Fq 'api(libs.jackson.databind)' "${contract_build}" && die \
    "contract still declares jackson-databind as api"

fixed_pattern_exists 'build/jacoco/test.exec' "${root_plugin}" && die \
    "root aggregated coverage still hardcodes test.exec"
fixed_pattern_exists ':engine:test' "${root_plugin}" && die "root coverage wiring still hardcodes module names"
fixed_pattern_exists ':protocol:test' "${root_plugin}" && die "root coverage wiring still hardcodes module names"
fixed_pattern_exists ':cli:test' "${root_plugin}" && die "root coverage wiring still hardcodes module names"
fixed_pattern_exists ':engine:jacocoTestCoverageVerification' "${root_plugin}" && die \
    "root coverage wiring still hardcodes module names"
fixed_pattern_exists ':protocol:jacocoTestCoverageVerification' "${root_plugin}" && die \
    "root coverage wiring still hardcodes module names"
fixed_pattern_exists ':cli:jacocoTestCoverageVerification' "${root_plugin}" && die \
    "root coverage wiring still hardcodes module names"
fixed_pattern_exists 'taskPathsByType(coverageSubprojects, Test::class.java)' "${root_plugin}" || die \
    "root aggregated coverage no longer discovers test tasks dynamically"
fixed_pattern_exists 'coverageSubprojects().flatMap { subproject ->' "${root_plugin}" || die \
    "root aggregated coverage no longer discovers JaCoCo execution data from all coverage subprojects"
fixed_pattern_exists 'testTask.extensions.getByType(JacocoTaskExtension::class.java).destinationFile' "${root_plugin}" || die \
    "root aggregated coverage no longer collects execution data from each module test task"

printf 'publication-contract regression: success\n'
