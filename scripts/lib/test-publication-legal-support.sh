#!/usr/bin/env bash
# Focused legal-surface assertions shared by the publication contract regression.

verify_publication_legal_contract() {
    local repository_root=$1
    local dockerfile_path=$2
    local container_workflow_path=$3
    local notice_path="${repository_root}/NOTICE"
    local wrapper_properties_path="${repository_root}/gradle/wrapper/gradle-wrapper.properties"
    local wrapper_jar_path="${repository_root}/gradle/wrapper/gradle-wrapper.jar"
    local custom_xml_test_path="${repository_root}/engine/src/test/resources/dev/erst/gridgrind/excel/CustomXMLMappings.xlsx"
    local custom_xml_example_path="${repository_root}/examples/custom-xml-assets/custom-xml-mapping.xlsx"
    local custom_xml_hash='6a5fbb160c7c9c2add5125ac07ceaf3f841d40e03d5c75025ab71946d49af289'
    local runtime_base_ref=''
    local wrapper_manifest=''

    grep -Fq 'provenance: mode=max' "${container_workflow_path}" || die \
        "container workflow does not publish explicit provenance"
    grep -Fq 'sbom: true' "${container_workflow_path}" || die \
        "container workflow does not publish an SBOM attestation"
    if grep -Fq 'org.opencontainers.image.licenses=' "${container_workflow_path}" \
        || grep -Fq 'org.opencontainers.image.licenses=' "${dockerfile_path}"; then
        die "container publication must not claim one incomplete aggregate license expression for all image contents"
    fi
    grep -Fq 'labels: ${{ steps.meta.outputs.labels }}' "${container_workflow_path}" && die \
        "container publication still applies metadata-action's repository-level MIT label to the aggregate image"
    grep -Fq 'COPY LICENSE NOTICE PATENTS.md LICENSE-APACHE-2.0 LICENSE-BSD-2-CLAUSE LICENSE-BSD-3-CLAUSE ./' "${dockerfile_path}" || die \
        "Docker builder stage does not receive the canonical legal files required by fat-JAR packaging"
    runtime_base_ref="$(awk '/^FROM azul\/zulu-openjdk:26-jre@sha256:/{print $2; exit}' "${dockerfile_path}")"
    [[ -n "${runtime_base_ref}" ]] || die \
        "Dockerfile no longer pins the published runtime base by digest"
    grep -Fq "${runtime_base_ref}" "${notice_path}" || die \
        "NOTICE does not identify the exact pinned published runtime base"
    grep -Fq 'GRIDGRIND_VERSION=${{ steps.target-tag.outputs.version }}' "${container_workflow_path}" || die \
        "container publication no longer supplies the GridGrind version build argument"

    grep -Fq 'gradle-9.7.1-bin.zip' "${wrapper_properties_path}" || die \
        "Gradle wrapper version changed without a source-distribution legal review"
    wrapper_manifest="$(unzip -p "${wrapper_jar_path}" META-INF/MANIFEST.MF | tr -d '\r')"
    grep -Fq 'SPDX-License-Identifier: Apache-2.0' <<<"${wrapper_manifest}" || die \
        "Gradle wrapper JAR no longer declares Apache-2.0"
    grep -Fq 'Gradle Wrapper 9.7.1' "${notice_path}" || die \
        "NOTICE omits the source-distributed Gradle wrapper"

    [[ "$(shasum -a 256 "${custom_xml_test_path}" | awk '{print $1}')" == "${custom_xml_hash}" ]] || die \
        "engine CustomXMLMappings.xlsx no longer matches its audited Apache POI provenance"
    [[ "$(shasum -a 256 "${custom_xml_example_path}" | awk '{print $1}')" == "${custom_xml_hash}" ]] || die \
        "example custom-xml-mapping.xlsx no longer matches its audited Apache POI provenance"
    grep -Fq "${custom_xml_hash}" "${notice_path}" || die \
        "NOTICE omits the audited Apache POI CustomXMLMappings.xlsx hash"
}
