#!/usr/bin/env bash
# Verify that the runtime image presents GridGrind and inherited legal material truthfully.

build_and_verify_image_legal_surface() {
    local image_ref=$1
    local repository_root=$2

    printf 'Docker smoke: building local image\n'
    docker_with_repo_config buildx build --load -t "${image_ref}" "${repository_root}" >/dev/null
    printf 'Docker smoke: verifying legal-material visibility\n'
    verify_image_legal_surface "${image_ref}"
}

verify_image_legal_surface() {
    local image_ref=$1
    local license_label=''
    local legal_file_name=''

    license_label="$(
        docker_with_repo_config image inspect "${image_ref}" \
            --format '{{index .Config.Labels "org.opencontainers.image.licenses"}}'
    )"
    [[ -z "${license_label}" || "${license_label}" == '<no value>' ]] || die \
        "Docker image publishes an incomplete aggregate OCI license expression: ${license_label}"

    [[ "$(docker_with_repo_config image inspect "${image_ref}" --format '{{index .Config.Labels "org.opencontainers.image.title"}}')" == 'GridGrind' ]] || die \
        "Docker image does not identify its packaged application title"

    for legal_file_name in \
        LICENSE \
        NOTICE \
        PATENTS.md \
        LICENSE-APACHE-2.0 \
        LICENSE-BSD-2-CLAUSE \
        LICENSE-BSD-3-CLAUSE; do
        docker_with_repo_config run --rm --entrypoint test "${image_ref}" \
            -f "/usr/share/doc/gridgrind/${legal_file_name}" || die \
            "Docker image is missing /usr/share/doc/gridgrind/${legal_file_name}"
    done

    docker_with_repo_config run --rm --entrypoint sh "${image_ref}" -c \
        'find /usr/lib/jvm -type f -name LICENSE -print -quit | grep -q . && test -f /usr/share/doc/base-files/copyright' || die \
        "Docker image does not retain the JRE and base-system legal material"
}
