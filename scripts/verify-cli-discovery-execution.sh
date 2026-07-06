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
import shutil
import subprocess
import sys
import time
import uuid
from pathlib import Path
from typing import Optional

repo_root = Path(sys.argv[1])
mode = sys.argv[2]
artifact_target = sys.argv[3]
temp_root = Path(sys.argv[4])
docker_run_user = sys.argv[5]
heartbeat_seconds = max(1, int(sys.argv[6]))
examples_root = repo_root / "examples"


def die(message: str) -> None:
    print(f"error: {message}", file=sys.stderr)
    raise SystemExit(1)


def progress(message: str) -> None:
    print(message, flush=True)


def launcher(command: list[str], cwd: Path) -> list[str]:
    if mode == "binary":
        return [artifact_target, *command]
    if mode == "jar":
        return ["java", "-jar", artifact_target, *command]
    if mode == "docker-image":
        docker_command = [
            "docker",
            "run",
            "--rm",
        ]
        if docker_run_user:
            docker_command.extend(["--user", docker_run_user])
        docker_command.extend(
            [
                "-v",
                f"{cwd}:/work",
                artifact_target,
                *command,
            ]
        )
        return [
            *docker_command,
        ]
    die(f"unsupported launcher mode {mode}")


def run(
    command: list[str],
    cwd: Path,
    progress_label: Optional[str] = None,
) -> subprocess.CompletedProcess[str]:
    logs_dir = temp_root / "_subprocess_logs"
    logs_dir.mkdir(parents=True, exist_ok=True)
    stdout_path = logs_dir / f"{uuid.uuid4()}-stdout.log"
    stderr_path = logs_dir / f"{uuid.uuid4()}-stderr.log"
    launch = launcher(command, cwd)
    started_at = time.monotonic()
    next_heartbeat_at = heartbeat_seconds
    with stdout_path.open("w", encoding="utf-8") as stdout_handle, stderr_path.open(
        "w",
        encoding="utf-8",
    ) as stderr_handle:
        process = subprocess.Popen(
            launch,
            cwd=cwd,
            text=True,
            stdout=stdout_handle,
            stderr=stderr_handle,
        )
        while True:
            returncode = process.poll()
            if returncode is not None:
                break
            elapsed_seconds = time.monotonic() - started_at
            if progress_label is not None and elapsed_seconds >= next_heartbeat_at:
                progress(
                    f"{progress_label} (still running after {int(elapsed_seconds)}s)"
                )
                next_heartbeat_at += heartbeat_seconds
            time.sleep(1)
    stdout = stdout_path.read_text(encoding="utf-8")
    stderr = stderr_path.read_text(encoding="utf-8")
    stdout_path.unlink(missing_ok=True)
    stderr_path.unlink(missing_ok=True)
    return subprocess.CompletedProcess(
        args=launch,
        returncode=returncode,
        stdout=stdout,
        stderr=stderr,
    )


def run_json(
    command: list[str],
    cwd: Path,
    progress_label: Optional[str] = None,
) -> object:
    completed = run(command, cwd, progress_label)
    if completed.returncode != 0:
        die(
            f"command failed ({completed.returncode}): {' '.join(launcher(command, cwd))}\n"
            + f"stdout: {completed.stdout}\n"
            + f"stderr: {completed.stderr}"
        )
    try:
        return json.loads(completed.stdout)
    except json.JSONDecodeError as exc:
        die(
            "command did not emit JSON: "
            + f"{' '.join(launcher(command, cwd))}\n{exc}\n{completed.stdout}"
        )


def artifact_path(path: Path, workspace: Path) -> str:
    if mode in {"binary", "jar"}:
        return str(path)
    if mode == "docker-image":
        return str(path.relative_to(workspace))
    die(f"unsupported launcher mode {mode}")


def copy_required_assets(request_dir: Path, required_paths: list[str]) -> None:
    for relative_path in required_paths:
        source_path = examples_root / relative_path
        target_path = request_dir / relative_path
        if not source_path.exists():
            die(f"missing published asset {source_path}")
        target_path.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source_path, target_path)


def execute_plan(
    kind: str,
    stable_id: str,
    ordinal: int,
    total: int,
    request_command: list[str],
    request_file_name: str,
    required_workspace_paths: list[str],
) -> None:
    workspace = temp_root / kind / stable_id.lower()
    workspace.mkdir(parents=True, exist_ok=True)
    request_path = workspace / request_file_name
    request_path.parent.mkdir(parents=True, exist_ok=True)
    doctor_path = workspace / "doctor.json"
    response_path = workspace / "response.json"

    progress(f"Discovery execution {kind} {ordinal}/{total}: {stable_id} printing request")
    printed = run(
        [*request_command, "--response", artifact_path(request_path, workspace)],
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

    progress(f"Discovery execution {kind} {ordinal}/{total}: {stable_id} copying required assets")
    copy_required_assets(request_path.parent, required_workspace_paths)

    progress(f"Discovery execution {kind} {ordinal}/{total}: {stable_id} doctoring request")
    doctor = run(
        [
            "--doctor-request",
            "--request",
            artifact_path(request_path, workspace),
            "--response",
            artifact_path(doctor_path, workspace),
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
    executed = run(
        [
            "--request",
            artifact_path(request_path, workspace),
            "--response",
            artifact_path(response_path, workspace),
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
recipe_catalog = run_json(
    ["--print-recipe-catalog"],
    catalog_workspace,
    "Discovery execution catalog: loading recipes",
)
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
        ["--print-recipe", "--lookup", example["id"]],
        example["requestFileName"],
        example["requiredWorkspacePaths"],
    )

for index, task_starter in enumerate(task_starter_entries, start=1):
    execute_plan(
        "task starters",
        task_starter["id"],
        index,
        len(task_starter_entries),
        ["--print-recipe", "--lookup", task_starter["id"]],
        task_starter["requestFileName"],
        task_starter["requiredWorkspacePaths"],
    )
PY

printf 'Verified CLI discovery execution surface via %s %s\n' "${mode}" "${target}"
