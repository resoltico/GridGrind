#!/usr/bin/env bash
# Keep the release-publication surface deterministic: pinned images, guarded publishing,
# attestations, accurate OCI labels, dynamic coverage wiring, and a narrow contract compile API.

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
readonly readme_file="${repo_root}/README.md"
readonly quick_start_doc="${repo_root}/docs/QUICK_START.md"
readonly root_plugin="${repo_root}/gradle/build-logic/src/main/kotlin/dev/erst/gridgrind/buildlogic/GridGrindRootConventionsPlugin.kt"
readonly contract_build="${repo_root}/contract/build.gradle.kts"
readonly cli_jar="${repo_root}/cli/build/libs/gridgrind.jar"
readonly docker_smoke_script="${repo_root}/scripts/docker-smoke.sh"
readonly docker_smoke_bind_mount_helper="${repo_root}/scripts/lib/docker-smoke-bind-mount-support.sh"
readonly container_verify_script="${repo_root}/scripts/verify-container-publication.sh"
readonly stage_contract_script="${repo_root}/scripts/check-stage-contract.sh"
readonly release_protocol_doc="${repo_root}/docs/RELEASE_PROTOCOL.md"
readonly temp_parent="${repo_root}/tmp/test-publication-contract"
source "${script_dir}/lib/test-publication-contract-diagnostic-support.sh"
test_root=''
jar_listing_path=''
archive_listing_path=''

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

dockerfile_copies_built_cli_jar() {
    local path=$1
    grep -Eq \
        '^COPY --from=build( --chown=[^[:space:]]+)? /workspace/cli/build/libs/gridgrind\.jar [^[:space:]]*gridgrind\.jar$' \
        "${path}"
}

grep -Eq \
    '^FROM azul/zulu-openjdk-alpine:26@sha256:[0-9a-f]{64} AS build$' \
    "${dockerfile}" || die "Dockerfile builder image is not digest-pinned"
grep -Eq \
    '^FROM azul/zulu-openjdk-alpine:26-jre@sha256:[0-9a-f]{64}$' \
    "${dockerfile}" || die "Dockerfile runtime image is not digest-pinned"

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
grep -Fq '!cli/build/libs/gridgrind.jar' "${dockerignore_file}" && die \
    ".dockerignore unexpectedly whitelists a prebuilt host JAR into the Docker build context"

command -v jar >/dev/null 2>&1 || die "jar is required for publication contract verification"
[[ -f "${cli_jar}" ]] || die "missing CLI fat JAR at ${cli_jar}"
mkdir -p "${temp_parent}"
test_root="${temp_parent}/run.$$"
rm -rf "${test_root}"
mkdir -p "${test_root}/archive-root/.codex/protocol" "${test_root}/archive-root/src"
jar_listing_path="${test_root}/cli-jar-listing.txt"
archive_listing_path="${test_root}/archive-listing.txt"

jar tf "${cli_jar}" > "${jar_listing_path}"
grep -Fq 'AGENTS.md' "${jar_listing_path}" && die \
    "CLI fat JAR unexpectedly contains /AGENTS.md"
grep -Fq '.codex/' "${jar_listing_path}" && die \
    "CLI fat JAR unexpectedly contains /.codex/"

verify_packaged_diagnostic_byte_stability "${cli_jar}" "${test_root}"

cp "${gitattributes_file}" "${test_root}/archive-root/.gitattributes"
printf '# synthetic agent entry point\n' > "${test_root}/archive-root/AGENTS.md"
printf '# synthetic codex doc\n' > "${test_root}/archive-root/.codex/AGENTS_EXTRA.md"
printf '# synthetic nested codex doc\n' > "${test_root}/archive-root/.codex/protocol/guide.md"
printf 'public file\n' > "${test_root}/archive-root/src/published.txt"
git -C "${test_root}/archive-root" -c init.defaultBranch=main init >/dev/null
git -C "${test_root}/archive-root" config user.name "GridGrind Test"
git -C "${test_root}/archive-root" config user.email "gridgrind-test@example.com"
git -C "${test_root}/archive-root" add .gitattributes AGENTS.md .codex/AGENTS_EXTRA.md .codex/protocol/guide.md src/published.txt
git -C "${test_root}/archive-root" commit -m "Archive surface fixture" >/dev/null
git -C "${test_root}/archive-root" archive --format=tar --output "${test_root}/archive.tar" HEAD
tar -tf "${test_root}/archive.tar" > "${archive_listing_path}"
grep -Fq 'src/published.txt' "${archive_listing_path}" || die \
    "git archive no longer includes ordinary tracked files for the public source asset"
grep -Fq 'AGENTS.md' "${archive_listing_path}" && die \
    "git archive still includes /AGENTS.md in the public source asset"
grep -Fq '.codex/AGENTS_EXTRA.md' "${archive_listing_path}" && die \
    "git archive still includes /.codex/AGENTS_EXTRA.md in the public source asset"
grep -Fq '.codex/protocol/guide.md' "${archive_listing_path}" && die \
    "git archive still includes nested /.codex content in the public source asset"

grep -Fq './scripts/verify-release-candidate-tag.sh "${{ steps.target-tag.outputs.tag }}"' \
    "${release_workflow}" || die "release workflow does not enforce the shared tag verifier"
grep -Fq './scripts/verify-release-candidate-tag.sh "${{ steps.target-tag.outputs.tag }}"' \
    "${container_workflow}" || die "container workflow does not enforce the shared tag verifier"
grep -Fq './scripts/verify-cli-contract.sh jar ./cli/build/libs/gridgrind.jar' \
    "${release_workflow}" || die "release workflow does not verify the packaged CLI contract"
grep -Fq 'COPY gradlew gradle.properties settings.gradle.kts build.gradle.kts ./' "${dockerfile}" || die \
    "Dockerfile no longer copies the Gradle entry files needed for the self-contained build"
grep -Fq 'COPY gradle ./gradle' "${dockerfile}" || die \
    "Dockerfile no longer copies the Gradle support directory into the builder stage"
grep -Fq 'COPY authoring-java ./authoring-java' "${dockerfile}" || die \
    "Dockerfile no longer copies authoring-java into the builder stage"
grep -Fq 'COPY cli ./cli' "${dockerfile}" || die \
    "Dockerfile no longer copies cli into the builder stage"
grep -Fq 'COPY contract ./contract' "${dockerfile}" || die \
    "Dockerfile no longer copies contract into the builder stage"
grep -Fq 'COPY engine ./engine' "${dockerfile}" || die \
    "Dockerfile no longer copies engine into the builder stage"
grep -Fq 'COPY excel-foundation ./excel-foundation' "${dockerfile}" || die \
    "Dockerfile no longer copies excel-foundation into the builder stage"
grep -Fq 'COPY executor ./executor' "${dockerfile}" || die \
    "Dockerfile no longer copies executor into the builder stage"
grep -Fq 'RUN --mount=type=cache,target=/root/.gradle ./gradlew --no-daemon :cli:shadowJar' "${dockerfile}" || die \
    "Dockerfile no longer builds the packaged CLI JAR inside the pinned builder stage"
dockerfile_copies_built_cli_jar "${dockerfile}" || die \
    "Dockerfile no longer copies the packaged CLI JAR from the builder stage into the runtime image"
grep -Fq 'COPY cli/build/libs/gridgrind.jar gridgrind.jar' "${dockerfile}" && die \
    "Dockerfile reintroduced a prebuilt host-JAR dependency"
grep -Fq 'docker buildx build --load -t gridgrind-local .' "${readme_file}" || die \
    "README.md no longer teaches the local repository Docker build path"
grep -Fq 'docker buildx build --load -t gridgrind-local .' "${quick_start_doc}" || die \
    "docs/QUICK_START.md no longer teaches the local repository Docker build path"
grep -Fq 'cli-shadow-jar-support.sh' "${docker_smoke_script}" && die \
    "docker smoke reintroduced the host-side CLI JAR helper dependency"
grep -Fq ':cli:shadowJar' "${docker_smoke_script}" && die \
    "docker smoke reintroduced a separate host-side CLI JAR rebuild"
grep -Fq 'docker_with_repo_config buildx build --load -t "${image_tag}" "${repo_root}" >/dev/null' \
    "${docker_smoke_script}" || die \
    "docker smoke no longer builds the repository-root Dockerfile through buildx --load"
grep -Fq '"${repo_root}/scripts/verify-cli-contract.sh" docker-image "${image_tag}"' \
    "${docker_smoke_script}" || die "docker smoke no longer verifies the local image CLI contract"
grep -Fq '"${repo_root}/scripts/verify-cli-discovery-execution.sh" docker-image "${image_tag}"' \
    "${docker_smoke_script}" || die \
    "docker smoke no longer executes the published examples and task starters from the local image"
grep -Fq 'source "${bind_mount_support}"' "${docker_smoke_script}" || die \
    "docker smoke no longer sources the bind-mount guidance helper"
grep -Fq 'verify_documented_bind_mount_user_guidance "${image_tag}" "${smoke_root}" "${docker_run_user}"' \
    "${docker_smoke_script}" || die \
    "docker smoke no longer calls the documented bind-mount guidance probe"
grep -Fq 'docker smoke no-user documented bind-mount run did not report IO_ERROR' \
    "${docker_smoke_bind_mount_helper}" || die \
    "bind-mount guidance helper no longer verifies the no-user failure path on non-remapped bind mounts"
grep -Fq 'docker smoke no-user documented bind-mount run completed both mounted writes despite failing' \
    "${docker_smoke_bind_mount_helper}" || die \
    "bind-mount guidance helper no longer proves the no-user failure path leaves the mounted writes incomplete"
grep -Fq 'docker smoke documented bind-mount run with --user did not write the workbook file' \
    "${docker_smoke_bind_mount_helper}" || die \
    "bind-mount guidance helper no longer verifies the documented command succeeds with --user"
grep -Fq "readonly streaming_read_request_rel='requests odd/request streaming readback [docker #smoke].json'" \
    "${docker_smoke_script}" || die "docker smoke no longer stages a separate streaming readback request"
grep -Fq -- '--request "${streaming_read_request_rel}"' "${docker_smoke_script}" || die \
    "docker smoke no longer reopens the streamed workbook for the summary readback"
grep -Fq '"level": "VERBOSE"' "${docker_smoke_script}" || die \
    "docker smoke no longer exercises verbose execution journaling from the packaged artifact"
grep -Fq 'docker smoke response did not include the structured execution journal' "${docker_smoke_script}" || die \
    "docker smoke no longer asserts response-journal presence"
grep -Fq 'docker smoke response did not retain VERBOSE journal events' "${docker_smoke_script}" || die \
    "docker smoke no longer asserts structured VERBOSE journal events"
grep -Fq 'docker smoke create request wrote unexpected stderr' "${docker_smoke_script}" || die \
    "docker smoke no longer asserts clean stderr for successful VERBOSE response-file writes"
grep -Fq 'run_verify_cli_contract "${image_name}:${expected_version}"' \
    "${container_verify_script}" || die "public container verification no longer checks the version tag contract"
grep -Fq 'run_verify_cli_contract "${image_name}:latest"' \
    "${container_verify_script}" || die "public container verification no longer checks the latest tag contract"
grep -Fq 'run_verify_cli_discovery_execution "${image_name}:${expected_version}"' \
    "${container_verify_script}" || die \
    "public container verification no longer executes published examples and task starters for the version tag"
grep -Fq 'run_verify_cli_discovery_execution "${image_name}:latest"' \
    "${container_verify_script}" || die \
    "public container verification no longer executes published examples and task starters for the latest tag"
grep -Fq 'scripts/test-verify-release-primary-checkout.sh' "${stage_contract_script}" || die \
    "Stage 4 contract no longer exercises the release primary-checkout regression"
grep -Fq './scripts/verify-release-primary-checkout.sh "$PRIMARY_CHECKOUT" "X.Y.Z"' \
    "${release_protocol_doc}" || die "release protocol no longer requires the primary-checkout closeout verifier"

grep -Fq 'checks: read' "${release_workflow}" || die "release workflow is missing checks: read permission"
grep -Fq 'checks: read' "${container_workflow}" || die "container workflow is missing checks: read permission"

grep -Fq 'provenance: mode=max' "${container_workflow}" || die "container workflow does not publish explicit provenance"
grep -Fq 'sbom: true' "${container_workflow}" || die "container workflow does not publish an SBOM attestation"
grep -Fqx '            org.opencontainers.image.licenses=MIT AND Apache-2.0 AND BSD-2-Clause AND BSD-3-Clause AND EDL-1.0' \
    "${container_workflow}" || die "container workflow still uses the wrong OCI license label"
grep -Fqx '            org.opencontainers.image.licenses=MIT' "${container_workflow}" && die \
    "container workflow still contains the legacy MIT-only OCI license label"
grep -Fqx '            org.opencontainers.image.licenses=MIT AND Apache-2.0 AND BSD-3-Clause' "${container_workflow}" && die \
    "container workflow still contains the incomplete 3-license OCI label (missing BSD-2-Clause and EDL-1.0)"

grep -Fq 'api(libs.jackson.annotations)' "${contract_build}" || die \
    "contract no longer exposes the Jackson annotations API needed by downstream compiles"
grep -Fq 'implementation(libs.jackson.databind)' "${contract_build}" || die \
    "contract no longer keeps jackson-databind off consumer compile classpaths"
grep -Fq 'api(libs.jackson.databind)' "${contract_build}" && die \
    "contract still declares jackson-databind as api"

fixed_pattern_exists 'build/jacoco/test.exec' "${root_plugin}" && die \
    "root aggregated coverage hardcodes test.exec"
fixed_pattern_exists ':engine:test' "${root_plugin}" && die "root coverage wiring hardcodes module names"
fixed_pattern_exists ':protocol:test' "${root_plugin}" && die "root coverage wiring hardcodes module names"
fixed_pattern_exists ':cli:test' "${root_plugin}" && die "root coverage wiring hardcodes module names"
fixed_pattern_exists ':engine:jacocoTestCoverageVerification' "${root_plugin}" && die \
    "root coverage wiring hardcodes module names"
fixed_pattern_exists ':protocol:jacocoTestCoverageVerification' "${root_plugin}" && die \
    "root coverage wiring hardcodes module names"
fixed_pattern_exists ':cli:jacocoTestCoverageVerification' "${root_plugin}" && die \
    "root coverage wiring hardcodes module names"
fixed_pattern_exists 'taskPathsByName(coverageSubprojects, "jacocoTestReport")' "${root_plugin}" || die \
    "root aggregated coverage no longer discovers module reports dynamically"
fixed_pattern_exists 'report.dependsOn(subprojectCoverageReports)' "${root_plugin}" || die \
    "root aggregated coverage no longer waits for dynamically discovered module reports"
fixed_pattern_exists 'coverageSubprojects().flatMap { subproject ->' "${root_plugin}" || die \
    "root aggregated coverage no longer discovers JaCoCo execution data from all coverage subprojects"
fixed_pattern_exists 'testTask.extensions.getByType(JacocoTaskExtension::class.java).destinationFile' "${root_plugin}" || die \
    "root aggregated coverage no longer collects execution data from each module test task"

printf 'publication-contract regression: success\n'
