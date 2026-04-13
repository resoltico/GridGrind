#!/usr/bin/env bash
# Verify that both the exact version tag and latest tag are publicly pullable and runnable from
# GHCR, accounting for short registry propagation delays after publication.

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

load_expected_description() {
    local repo_root_path=$1
    local description

    description="$(
        awk -F= '
            $1 == "gridgrindDescription" {
                sub(/^[^=]*=/, "", $0)
                print $0
                exit
            }
        ' "${repo_root_path}/gradle.properties"
    )"
    [[ -n "${description}" ]] || die "unable to load gridgrindDescription from ${repo_root_path}/gradle.properties"
    printf '%s' "${description}"
}

readonly image_name="${1:-}"
readonly expected_version="${2:-}"
readonly retry_count="${GRIDGRIND_PUBLICATION_VERIFY_RETRIES:-12}"
readonly retry_delay_seconds="${GRIDGRIND_PUBLICATION_VERIFY_DELAY_SECONDS:-10}"
readonly script_dir="$(resolve_script_dir)"
readonly repo_root="$(cd -P -- "${script_dir}/.." && pwd)"
readonly expected_description="$(load_expected_description "${repo_root}")"
docker_config_dir=''

[[ -n "${image_name}" ]] || die "image name is required"
[[ -n "${expected_version}" ]] || die "expected version is required"

expected_output="$(printf 'GridGrind %s\n%s' "${expected_version}" "${expected_description}")"

cleanup() {
    [[ -n "${docker_config_dir}" ]] || return
    rm -rf "${docker_config_dir}"
}

trap cleanup EXIT

docker_config_dir="$(mktemp -d)"

anonymous_docker() {
    docker --config "${docker_config_dir}" "$@"
}

verify_ref() {
    local tag_ref=$1
    local image_ref="${image_name}:${tag_ref}"
    local attempt version_output

    for ((attempt = 1; attempt <= retry_count; attempt++)); do
        if anonymous_docker pull "${image_ref}" >/dev/null 2>&1; then
            version_output="$(anonymous_docker run --rm "${image_ref}" --version 2>/dev/null | tr -d '\r')"
            if [[ "${version_output}" == "${expected_output}" ]]; then
                printf 'Verified published container: %s\n' "${image_ref}"
                return
            fi
        fi

        if (( attempt < retry_count )); then
            sleep "${retry_delay_seconds}"
        fi
    done

    die "published container ${image_ref} did not report ${expected_output}"
}

verify_ref "${expected_version}"
verify_ref latest
