#!/usr/bin/env bash
# Execute every published built-in example and task starter from a packaged CLI artifact.

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
readonly temp_parent="${repo_root}/tmp/verify-cli-discovery-execution"
python3_path="$(command -v python3 || true)"
[[ -n "${python3_path}" ]] || die "python3 is required for discovery execution verification"
readonly heartbeat_seconds="${GRIDGRIND_DISCOVERY_EXECUTION_HEARTBEAT_SECONDS:-20}"

# shellcheck source=/dev/null
source "${repo_root}/scripts/lib/cli-shadow-jar-support.sh"

mode='jar'
target=''
docker_run_user="${GRIDGRIND_DOCKER_RUN_USER:-}"
if [[ $# -eq 0 ]]; then
    target="$(ensure_cli_shadow_jar "${repo_root}")"
elif [[ $# -eq 1 ]]; then
    case "${1}" in
        binary)
            die "binary mode requires an executable target"
            ;;
        jar)
            target="$(ensure_cli_shadow_jar "${repo_root}")"
            ;;
        docker-image)
            die "docker-image mode requires an image reference"
            ;;
        *)
            mode='binary'
            target="$(cd -P -- "$(dirname -- "${1}")" && pwd)/$(basename -- "${1}")"
            ;;
    esac
elif [[ $# -eq 2 ]]; then
    mode="${1}"
    target="${2}"
else
    die "usage: $0 [jar <path>|docker-image <image-ref>|<jar-path>]"
fi

case "${mode}" in
    binary)
        target="$(cd -P -- "$(dirname -- "${target}")" && pwd)/$(basename -- "${target}")"
        [[ -x "${target}" ]] || die "missing executable CLI launcher: ${target}"
        ;;
    jar)
        command -v java >/dev/null 2>&1 || die "java is required for jar verification"
        target="$(cd -P -- "$(dirname -- "${target}")" && pwd)/$(basename -- "${target}")"
        [[ -f "${target}" ]] || die "missing CLI jar: ${target}"
        ;;
    docker-image)
        command -v docker >/dev/null 2>&1 || die "docker is required for docker-image verification"
        [[ -n "${target}" ]] || die "docker-image mode requires an image reference"
        if [[ -z "${docker_run_user}" ]] && command -v id >/dev/null 2>&1; then
            docker_run_user="$(id -u):$(id -g)"
        fi
        ;;
    *)
        die "unsupported mode ${mode}; expected binary, jar, or docker-image"
        ;;
esac

mkdir -p "${temp_parent}"
temp_dir="$(mktemp -d "${temp_parent%/}/run.XXXXXX")"
cleanup() {
    rm -rf "${temp_dir}"
}
trap cleanup EXIT

"${python3_path}" - \
    "${repo_root}" \
    "${mode}" \
    "${target}" \
    "${temp_dir}" \
    "${docker_run_user}" \
    "${heartbeat_seconds}" <<'PY'
import json
import subprocess
import sys
from pathlib import Path

repo_root = Path(sys.argv[1])
mode = sys.argv[2]
artifact_target = sys.argv[3]
temp_root = Path(sys.argv[4])
docker_run_user = sys.argv[5]
heartbeat_seconds = max(1, int(sys.argv[6]))
sys.path.insert(0, str(repo_root / "scripts" / "lib"))
from discovery_execution_support import ArtifactRunner

def die(message: str) -> None:
    print(f"error: {message}", file=sys.stderr)
    raise SystemExit(1)

def progress(message: str) -> None:
    print(message, flush=True)

runner = ArtifactRunner(
    mode, artifact_target, temp_root, docker_run_user, heartbeat_seconds, progress
)

def execute_plan(
    kind: str,
    stable_id: str,
    ordinal: int,
    total: int,
    request_file_name: str,
    required_workspace_paths: list[str],
) -> None:
    workspace_parent = temp_root / kind
    workspace_parent.mkdir(parents=True, exist_ok=True)
    workspace = workspace_parent / stable_id.lower()
    request_path = workspace / request_file_name
    if required_workspace_paths:
        progress(f"Discovery execution {kind} {ordinal}/{total}: {stable_id} materializing recipe workspace")
        materialized = runner.run(
            [
                "--materialize-recipe",
                "--lookup",
                stable_id,
                "--workspace",
                runner.artifact_path(workspace, workspace_parent),
            ],
            workspace_parent,
            f"Discovery execution {kind} {ordinal}/{total}: {stable_id} materializing recipe workspace",
        )
        if materialized.returncode != 0:
            die(
                f"{kind} {stable_id} did not materialize successfully\n"
                + f"stdout: {materialized.stdout}\n"
                + f"stderr: {materialized.stderr}"
            )
        if json.loads(materialized.stdout) != json.loads(request_path.read_text()):
            die(f"{kind} {stable_id} materialized request differs from its primary output")
    else:
        workspace.mkdir(parents=True, exist_ok=True)
        progress(f"Discovery execution {kind} {ordinal}/{total}: {stable_id} printing request")
        printed = runner.run(
            [
                "--print-recipe",
                "--lookup",
                stable_id,
                "--response",
                runner.artifact_path(request_path, workspace),
            ],
            workspace,
            f"Discovery execution {kind} {ordinal}/{total}: {stable_id} printing request",
        )
        if printed.returncode != 0:
            die(
                f"{kind} {stable_id} did not print successfully\n"
                + f"stdout: {printed.stdout}\n"
                + f"stderr: {printed.stderr}"
            )
    if not request_path.exists():
        die(f"{kind} {stable_id} did not create request file {request_path}")

    progress(f"Discovery execution {kind} {ordinal}/{total}: {stable_id} preparing workspace")
    grant_path = workspace / "host-grant.json"
    grant_working_directory_arguments = (
        ["--grant-working-directory", str(workspace)]
        if mode == "docker-image"
        else []
    )
    prepared = subprocess.run(
        [
            sys.executable,
            str(repo_root / "scripts" / "write-fixture-grant.py"),
            "--request",
            str(request_path),
            "--output",
            str(grant_path),
            *grant_working_directory_arguments,
            "--prepare-save-as-parent",
        ],
        capture_output=True,
        encoding="utf-8",
    )
    if prepared.returncode != 0:
        die(f"could not prepare {kind} {stable_id}: {prepared.stderr.strip()}")
    doctor_path = workspace / "doctor.json"
    response_path = workspace / "response.json"
    secrets_provider = runner.fixture_secret_provider(stable_id, workspace)
    secret_provider_arguments = (
        ["--secrets-provider", runner.artifact_path(secrets_provider, workspace)]
        if secrets_provider is not None
        else []
    )

    progress(f"Discovery execution {kind} {ordinal}/{total}: {stable_id} doctoring request")
    doctor = runner.run(
        [
            "--doctor-request",
            "--request",
            runner.artifact_path(request_path, workspace),
            "--grant",
            runner.artifact_path(grant_path, workspace),
            *secret_provider_arguments,
            "--response",
            runner.artifact_path(doctor_path, workspace),
        ],
        workspace,
        f"Discovery execution {kind} {ordinal}/{total}: {stable_id} doctoring request",
    )
    if doctor.returncode != 0:
        die(
            f"{kind} {stable_id} did not doctor cleanly\n"
            + f"stdout: {doctor.stdout}\n"
            + f"stderr: {doctor.stderr}\n"
            + f"doctor report: {doctor_path.read_text() if doctor_path.exists() else '<missing>'}"
        )
    doctor_report = json.loads(doctor_path.read_text())
    if doctor_report.get("valid") is not True:
        die(f"{kind} {stable_id} doctor report was not valid: {doctor_report}")

    progress(f"Discovery execution {kind} {ordinal}/{total}: {stable_id} executing request")
    executed = runner.run(
        [
            "--request",
            runner.artifact_path(request_path, workspace),
            "--grant",
            runner.artifact_path(grant_path, workspace),
            *secret_provider_arguments,
            "--response",
            runner.artifact_path(response_path, workspace),
        ],
        workspace,
        f"Discovery execution {kind} {ordinal}/{total}: {stable_id} executing request",
    )
    if executed.returncode != 0:
        die(
            f"{kind} {stable_id} did not execute successfully\n"
            + f"stdout: {executed.stdout}\n"
            + f"stderr: {executed.stderr}\n"
            + f"response: {response_path.read_text() if response_path.exists() else '<missing>'}"
        )
    response = json.loads(response_path.read_text())
    if "problem" in response:
        die(f"{kind} {stable_id} returned a failure response: {response}")
    progress(f"Discovery execution {kind} {ordinal}/{total}: {stable_id} succeeded")

catalog_workspace = temp_root / "_catalog"
catalog_workspace.mkdir(parents=True, exist_ok=True)
try:
    recipe_catalog = runner.run_json(
        ["--print-recipe-catalog"],
        catalog_workspace,
        "Discovery execution catalog: loading recipes",
    )
except RuntimeError as exception:
    die(str(exception))
recipe_entries = recipe_catalog["recipes"]
example_entries = [
    recipe for recipe in recipe_entries if recipe.get("view") == "EXAMPLE"
]
task_starter_entries = [
    recipe for recipe in recipe_entries if recipe.get("view") == "TASK_STARTER"
]

for index, example in enumerate(example_entries, start=1):
    execute_plan(
        "examples",
        example["id"],
        index,
        len(example_entries),
        example["requestFileName"],
        example["requiredWorkspacePaths"],
    )

for index, task_starter in enumerate(task_starter_entries, start=1):
    execute_plan(
        "task starters",
        task_starter["id"],
        index,
        len(task_starter_entries),
        task_starter["requestFileName"],
        task_starter["requiredWorkspacePaths"],
    )
PY
printf 'Verified CLI discovery execution surface via %s %s\n' "${mode}" "${target}"
