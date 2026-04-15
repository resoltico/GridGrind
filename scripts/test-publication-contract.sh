#!/usr/bin/env bash
# Keep the release-publication surface deterministic: pinned base image, guarded workflow-dispatch
# publishing, explicit attestations, accurate OCI labels, dynamic root coverage wiring, and a
# narrow protocol compile API.

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
readonly dockerfile="${repo_root}/Dockerfile"
readonly release_workflow="${repo_root}/.github/workflows/release.yml"
readonly container_workflow="${repo_root}/.github/workflows/container.yml"
readonly root_plugin="${repo_root}/gradle/build-logic/src/main/kotlin/dev/erst/gridgrind/buildlogic/GridGrindRootConventionsPlugin.kt"
readonly protocol_build="${repo_root}/protocol/build.gradle.kts"
readonly docker_smoke_script="${repo_root}/scripts/docker-smoke.sh"
readonly container_verify_script="${repo_root}/scripts/verify-container-publication.sh"

grep -Eq \
    '^FROM azul/zulu-openjdk-alpine:26-jre@sha256:[0-9a-f]{64}$' \
    "${dockerfile}" || die "Dockerfile base image is not digest-pinned"

grep -Fq './scripts/verify-release-candidate-tag.sh "${{ steps.target-tag.outputs.tag }}"' \
    "${release_workflow}" || die "release workflow does not enforce the shared tag verifier"
grep -Fq './scripts/verify-release-candidate-tag.sh "${{ steps.target-tag.outputs.tag }}"' \
    "${container_workflow}" || die "container workflow does not enforce the shared tag verifier"
grep -Fq './scripts/verify-cli-contract.sh jar ./cli/build/libs/gridgrind.jar' \
    "${release_workflow}" || die "release workflow does not verify the packaged CLI contract"
grep -Fq '"${repo_root}/scripts/verify-cli-contract.sh" docker-image "${image_tag}"' \
    "${docker_smoke_script}" || die "docker smoke no longer verifies the local image CLI contract"
grep -Fq '"${verify_cli_contract_script}" docker-image "${image_name}:${expected_version}"' \
    "${container_verify_script}" || die "public container verification no longer checks the version tag contract"
grep -Fq '"${verify_cli_contract_script}" docker-image "${image_name}:latest"' \
    "${container_verify_script}" || die "public container verification no longer checks the latest tag contract"

grep -Fq 'checks: read' "${release_workflow}" || die "release workflow is missing checks: read permission"
grep -Fq 'checks: read' "${container_workflow}" || die "container workflow is missing checks: read permission"

grep -Fq 'provenance: mode=max' "${container_workflow}" || die "container workflow does not publish explicit provenance"
grep -Fq 'sbom: true' "${container_workflow}" || die "container workflow does not publish an SBOM attestation"
grep -Fqx '            org.opencontainers.image.licenses=MIT AND Apache-2.0 AND BSD-3-Clause' \
    "${container_workflow}" || die "container workflow still uses the wrong OCI license label"
grep -Fqx '            org.opencontainers.image.licenses=MIT' "${container_workflow}" && die \
    "container workflow still contains the legacy MIT-only OCI license label"

grep -Fq 'api(libs.jackson.annotations)' "${protocol_build}" || die \
    "protocol no longer exposes the Jackson annotations API needed by downstream compiles"
grep -Fq 'implementation(libs.jackson.databind)' "${protocol_build}" || die \
    "protocol still leaks jackson-databind on consumer compile classpaths"
grep -Fq 'api(libs.jackson.databind)' "${protocol_build}" && die \
    "protocol still declares jackson-databind as api"

rg -Fq 'build/jacoco/test.exec' "${root_plugin}" && die \
    "root aggregated coverage still hardcodes test.exec"
rg -Fq ':engine:test' "${root_plugin}" && die "root coverage wiring still hardcodes module names"
rg -Fq ':protocol:test' "${root_plugin}" && die "root coverage wiring still hardcodes module names"
rg -Fq ':cli:test' "${root_plugin}" && die "root coverage wiring still hardcodes module names"
rg -Fq ':engine:jacocoTestCoverageVerification' "${root_plugin}" && die \
    "root coverage wiring still hardcodes module names"
rg -Fq ':protocol:jacocoTestCoverageVerification' "${root_plugin}" && die \
    "root coverage wiring still hardcodes module names"
rg -Fq ':cli:jacocoTestCoverageVerification' "${root_plugin}" && die \
    "root coverage wiring still hardcodes module names"
rg -Fq 'include("*.exec")' "${root_plugin}" || die \
    "root aggregated coverage no longer collects all module .exec files"
rg -Fq 'taskPathsByType(coverageSubprojects, Test::class.java)' "${root_plugin}" || die \
    "root aggregated coverage no longer discovers test tasks dynamically"

printf 'publication-contract regression: success\n'
