#!/usr/bin/env bash
# Verify the packaged installShadowDist launcher surface, not just the fat JAR surface.

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
readonly gradlew="${repo_root}/gradlew"
readonly verify_script="${repo_root}/scripts/verify-cli-contract.sh"
readonly install_root="${repo_root}/cli/build/install/cli-shadow/bin"
readonly generated_scripts_root="${repo_root}/cli/build/scriptsShadow"
readonly packaged_launcher="${install_root}/gridgrind"
readonly legacy_launcher="${install_root}/cli"
readonly legacy_windows_launcher="${install_root}/cli.bat"
readonly legacy_generated_launcher="${generated_scripts_root}/cli"
readonly legacy_generated_windows_launcher="${generated_scripts_root}/cli.bat"

mkdir -p "${install_root}" "${generated_scripts_root}"
printf '#!/usr/bin/env bash\nexit 99\n' > "${legacy_launcher}"
printf '@echo off\r\nexit /b 99\r\n' > "${legacy_windows_launcher}"
printf '#!/usr/bin/env bash\nexit 99\n' > "${legacy_generated_launcher}"
printf '@echo off\r\nexit /b 99\r\n' > "${legacy_generated_windows_launcher}"
chmod +x "${legacy_launcher}" "${legacy_generated_launcher}"

"${gradlew}" :cli:installShadowDist --console=plain --no-daemon >/dev/null

[[ -x "${packaged_launcher}" ]] || die \
    "installShadowDist did not produce the packaged gridgrind launcher at ${packaged_launcher}"
[[ ! -e "${legacy_launcher}" ]] || die \
    "installShadowDist still produced the misleading legacy launcher name at ${legacy_launcher}"
[[ ! -e "${legacy_windows_launcher}" ]] || die \
    "installShadowDist still produced the misleading legacy Windows launcher name at ${legacy_windows_launcher}"
[[ ! -e "${legacy_generated_launcher}" ]] || die \
    "startShadowScripts left the stale legacy launcher behind at ${legacy_generated_launcher}"
[[ ! -e "${legacy_generated_windows_launcher}" ]] || die \
    "startShadowScripts left the stale legacy Windows launcher behind at ${legacy_generated_windows_launcher}"

"${verify_script}" binary "${packaged_launcher}" >/dev/null

printf 'cli-distribution-surface regression: success\n'
