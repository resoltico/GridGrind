#!/usr/bin/env bash
# Shared helper for locating and materializing the packaged CLI fat JAR.

resolve_cli_shadow_jar_path() {
    local helper_repo_root=$1
    printf '%s\n' "${helper_repo_root}/cli/build/libs/gridgrind.jar"
}

ensure_cli_shadow_jar() {
    local helper_repo_root=$1
    local jar_path
    jar_path="$(resolve_cli_shadow_jar_path "${helper_repo_root}")"
    "${helper_repo_root}/gradlew" --no-daemon :cli:shadowJar --rerun --console=plain >&2
    [[ -f "${jar_path}" ]] || {
        printf 'error: expected CLI shadow jar at %s after build\n' "${jar_path}" >&2
        return 1
    }
    printf '%s\n' "${jar_path}"
}
