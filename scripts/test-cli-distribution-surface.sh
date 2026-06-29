#!/usr/bin/env bash
# Verify the packaged launcher and archive surfaces, not just the fat JAR surface.

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
readonly thin_install_root="${repo_root}/cli/build/install/cli/bin"
readonly install_root="${repo_root}/cli/build/install/gridgrind/bin"
readonly legacy_install_root="${repo_root}/cli/build/install/cli-shadow"
readonly generated_scripts_root="${repo_root}/cli/build/scriptsShadow"
readonly distribution_root="${repo_root}/cli/build/distributions"
readonly version="$(awk -F= '/^version=/{print $2}' "${repo_root}/gradle.properties")"
readonly thin_launcher="${thin_install_root}/cli"
readonly duplicate_thin_launcher="${thin_install_root}/gridgrind"
readonly duplicate_thin_windows_launcher="${thin_install_root}/gridgrind.bat"
readonly packaged_launcher="${install_root}/gridgrind"
readonly old_named_launcher="${install_root}/cli"
readonly old_named_windows_launcher="${install_root}/cli.bat"
readonly old_named_generated_launcher="${generated_scripts_root}/cli"
readonly old_named_generated_windows_launcher="${generated_scripts_root}/cli.bat"
readonly packaged_zip="${distribution_root}/gridgrind-${version}.zip"
readonly packaged_tar="${distribution_root}/gridgrind-${version}.tar"
readonly legacy_packaged_zip="${distribution_root}/cli-shadow-${version}.zip"
readonly legacy_packaged_tar="${distribution_root}/cli-shadow-${version}.tar"
readonly stale_zip="${distribution_root}/gridgrind-0.00.0.zip"
readonly stale_tar="${distribution_root}/gridgrind-0.00.0.tar"

mkdir -p "${thin_install_root}" "${install_root}" "${legacy_install_root}/bin" "${generated_scripts_root}" "${distribution_root}"
printf '#!/usr/bin/env bash\nexit 99\n' > "${duplicate_thin_launcher}"
printf '@echo off\r\nexit /b 99\r\n' > "${duplicate_thin_windows_launcher}"
printf '#!/usr/bin/env bash\nexit 99\n' > "${old_named_launcher}"
printf '@echo off\r\nexit /b 99\r\n' > "${old_named_windows_launcher}"
printf '#!/usr/bin/env bash\nexit 99\n' > "${old_named_generated_launcher}"
printf '@echo off\r\nexit /b 99\r\n' > "${old_named_generated_windows_launcher}"
printf '#!/usr/bin/env bash\nexit 99\n' > "${legacy_install_root}/bin/gridgrind"
printf 'stale zip\n' > "${stale_zip}"
printf 'stale tar\n' > "${stale_tar}"
printf 'legacy zip\n' > "${legacy_packaged_zip}"
printf 'legacy tar\n' > "${legacy_packaged_tar}"
chmod +x "${duplicate_thin_launcher}" "${old_named_launcher}" "${old_named_generated_launcher}" "${legacy_install_root}/bin/gridgrind"

"${gradlew}" :cli:installDist --console=plain --no-daemon >/dev/null
"${gradlew}" :cli:installShadowDist --console=plain --no-daemon >/dev/null
"${gradlew}" :cli:shadowDistZip :cli:shadowDistTar --console=plain --no-daemon >/dev/null

[[ -x "${thin_launcher}" ]] || die \
    "installDist did not produce the renamed thin launcher at ${thin_launcher}"
[[ ! -e "${duplicate_thin_launcher}" ]] || die \
    "installDist still produced the duplicate gridgrind launcher at ${duplicate_thin_launcher}"
[[ ! -e "${duplicate_thin_windows_launcher}" ]] || die \
    "installDist still produced the duplicate Windows gridgrind launcher at ${duplicate_thin_windows_launcher}"
[[ -x "${packaged_launcher}" ]] || die \
    "installShadowDist did not produce the packaged gridgrind launcher at ${packaged_launcher}"
[[ ! -e "${legacy_install_root}" ]] || die \
    "installShadowDist left the legacy install tree behind at ${legacy_install_root}"
[[ ! -e "${old_named_launcher}" ]] || die \
    "installShadowDist produced the old launcher name at ${old_named_launcher}"
[[ ! -e "${old_named_windows_launcher}" ]] || die \
    "installShadowDist produced the old Windows launcher name at ${old_named_windows_launcher}"
[[ ! -e "${old_named_generated_launcher}" ]] || die \
    "startShadowScripts left the stale old launcher behind at ${old_named_generated_launcher}"
[[ ! -e "${old_named_generated_windows_launcher}" ]] || die \
    "startShadowScripts left the stale old Windows launcher behind at ${old_named_generated_windows_launcher}"
[[ -f "${packaged_zip}" ]] || die \
    "shadowDistZip did not produce the packaged archive at ${packaged_zip}"
[[ -f "${packaged_tar}" ]] || die \
    "shadowDistTar did not produce the packaged archive at ${packaged_tar}"
[[ ! -e "${legacy_packaged_zip}" ]] || die \
    "shadowDistZip left the legacy archive name behind at ${legacy_packaged_zip}"
[[ ! -e "${legacy_packaged_tar}" ]] || die \
    "shadowDistTar left the legacy archive name behind at ${legacy_packaged_tar}"
[[ ! -e "${stale_zip}" ]] || die \
    "shadowDistZip left the stale archive behind at ${stale_zip}"
[[ ! -e "${stale_tar}" ]] || die \
    "shadowDistTar left the stale archive behind at ${stale_tar}"

"${verify_script}" binary "${packaged_launcher}" >/dev/null

printf 'cli-distribution-surface regression: success\n'
