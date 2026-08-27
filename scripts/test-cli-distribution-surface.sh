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
readonly legacy_thin_install_root="${repo_root}/cli/build/install/cli"
readonly install_root="${repo_root}/cli/build/install/gridgrind/bin"
readonly legacy_install_root="${repo_root}/cli/build/install/cli-shadow"
readonly legacy_start_scripts_root="${repo_root}/cli/build/scripts"
readonly generated_scripts_root="${repo_root}/cli/build/scriptsShadow"
readonly distribution_root="${repo_root}/cli/build/distributions"
readonly version="$(awk -F= '/^version=/{print $2}' "${repo_root}/gradle.properties")"
readonly packaged_launcher="${install_root}/gridgrind"
readonly old_named_launcher="${install_root}/cli"
readonly old_named_windows_launcher="${install_root}/cli.bat"
readonly old_named_legacy_launcher="${legacy_start_scripts_root}/cli"
readonly old_named_legacy_windows_launcher="${legacy_start_scripts_root}/cli.bat"
readonly old_named_generated_launcher="${generated_scripts_root}/cli"
readonly old_named_generated_windows_launcher="${generated_scripts_root}/cli.bat"
readonly packaged_zip="${distribution_root}/gridgrind-${version}.zip"
readonly packaged_tar="${distribution_root}/gridgrind-${version}.tar"
readonly legacy_packaged_zip="${distribution_root}/cli-shadow-${version}.zip"
readonly legacy_packaged_tar="${distribution_root}/cli-shadow-${version}.tar"
readonly stale_zip="${distribution_root}/gridgrind-0.00.0.zip"
readonly stale_tar="${distribution_root}/gridgrind-0.00.0.tar"
readonly version_fallback_response_path="${distribution_root}/version-fallback-existing.json"
readonly version_primary_stdout_path="${distribution_root}/version-primary.stdout"
readonly version_primary_stderr_path="${distribution_root}/version-primary.stderr"
readonly version_fallback_stdout_path="${distribution_root}/version-fallback.stdout"
readonly version_fallback_stderr_path="${distribution_root}/version-fallback.stderr"

cleanup() {
    rm -f \
        "${version_fallback_response_path}" \
        "${version_primary_stdout_path}" \
        "${version_primary_stderr_path}" \
        "${version_fallback_stdout_path}" \
        "${version_fallback_stderr_path}"
    rm -rf \
        "${legacy_thin_install_root}" \
        "${legacy_start_scripts_root}" \
        "${legacy_install_root}"
}

trap cleanup EXIT
cleanup

mkdir -p "${install_root}" "${generated_scripts_root}" "${distribution_root}"
printf '#!/usr/bin/env bash\nexit 99\n' > "${old_named_launcher}"
printf '@echo off\r\nexit /b 99\r\n' > "${old_named_windows_launcher}"
printf '#!/usr/bin/env bash\nexit 99\n' > "${old_named_generated_launcher}"
printf '@echo off\r\nexit /b 99\r\n' > "${old_named_generated_windows_launcher}"
mkdir -p "${legacy_install_root}/bin"
printf '#!/usr/bin/env bash\nexit 99\n' > "${legacy_install_root}/bin/gridgrind"
printf 'stale zip\n' > "${stale_zip}"
printf 'stale tar\n' > "${stale_tar}"
printf 'legacy zip\n' > "${legacy_packaged_zip}"
printf 'legacy tar\n' > "${legacy_packaged_tar}"
chmod +x \
    "${old_named_launcher}" \
    "${old_named_generated_launcher}" \
    "${legacy_install_root}/bin/gridgrind"

"${gradlew}" :cli:installDist --console=plain --no-daemon >/dev/null
[[ ! -e "${legacy_thin_install_root}" ]] || die \
    "installDist unexpectedly materialized the retired thin install tree at ${legacy_thin_install_root}"
[[ ! -e "${legacy_start_scripts_root}" ]] || die \
    "installDist unexpectedly materialized the retired thin start scripts at ${legacy_start_scripts_root}"

mkdir -p "${legacy_thin_install_root}/bin" "${legacy_start_scripts_root}"
printf '#!/usr/bin/env bash\nexit 99\n' > "${legacy_thin_install_root}/bin/cli"
printf '@echo off\r\nexit /b 99\r\n' > "${legacy_thin_install_root}/bin/cli.bat"
printf '#!/usr/bin/env bash\nexit 99\n' > "${old_named_legacy_launcher}"
printf '@echo off\r\nexit /b 99\r\n' > "${old_named_legacy_windows_launcher}"
"${gradlew}" :cli:installShadowDist --console=plain --no-daemon >/dev/null
"${gradlew}" :cli:shadowDistZip :cli:shadowDistTar --console=plain --no-daemon >/dev/null

[[ -x "${packaged_launcher}" ]] || die \
    "installShadowDist did not produce the packaged gridgrind launcher at ${packaged_launcher}"
[[ ! -e "${legacy_thin_install_root}" ]] || die \
    "installShadowDist left the retired thin install tree behind at ${legacy_thin_install_root}"
[[ ! -e "${legacy_install_root}" ]] || die \
    "installShadowDist left the legacy install tree behind at ${legacy_install_root}"
[[ ! -e "${legacy_start_scripts_root}" ]] || die \
    "installShadowDist left the retired thin start-script tree behind at ${legacy_start_scripts_root}"
[[ ! -e "${old_named_launcher}" ]] || die \
    "installShadowDist produced the old launcher name at ${old_named_launcher}"
[[ ! -e "${old_named_windows_launcher}" ]] || die \
    "installShadowDist produced the old Windows launcher name at ${old_named_windows_launcher}"
[[ ! -e "${old_named_legacy_launcher}" ]] || die \
    "installShadowDist left the retired thin start-script launcher behind at ${old_named_legacy_launcher}"
[[ ! -e "${old_named_legacy_windows_launcher}" ]] || die \
    "installShadowDist left the retired thin Windows launcher behind at ${old_named_legacy_windows_launcher}"
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

"${packaged_launcher}" --version \
    > "${version_primary_stdout_path}" \
    2> "${version_primary_stderr_path}"
[[ ! -s "${version_primary_stderr_path}" ]] || die \
    "packaged gridgrind --version wrote unexpected stderr without --response"

printf 'sentinel\n' > "${version_fallback_response_path}"
set +e
"${packaged_launcher}" --version --response "${version_fallback_response_path}" \
    > "${version_fallback_stdout_path}" \
    2> "${version_fallback_stderr_path}"
version_fallback_exit_code=$?
set -e

[[ ${version_fallback_exit_code} -eq 1 ]] || die \
    "packaged gridgrind --version --response existing-file exited ${version_fallback_exit_code} instead of 1"
grep -Fqx 'sentinel' "${version_fallback_response_path}" || die \
    "packaged gridgrind --version overwrote the pre-existing response target"
cmp -s "${version_primary_stdout_path}" "${version_fallback_stdout_path}" || die \
    "packaged gridgrind --version fallback stdout no longer preserves the primary payload"
grep -Eq '"wroteTo"[[:space:]]*:[[:space:]]*"STDOUT"' "${version_fallback_stderr_path}" || die \
    "packaged gridgrind --version fallback stderr no longer identifies stdout fallback transport"
python3 - "${version_fallback_stderr_path}" "${version_fallback_response_path}" <<'PY'
import json
import sys
from pathlib import Path

transport = json.loads(Path(sys.argv[1]).read_text())
expected = {
    "reason": "RESPONSE_WRITE_FAILED",
    "wroteTo": "STDOUT",
    "responsePath": sys.argv[2],
}
if transport != expected:
    print(
        "error: packaged gridgrind --version fallback stderr is not the exact transport notice",
        file=sys.stderr,
    )
    raise SystemExit(1)
PY

printf 'cli-distribution-surface regression: success\n'
