"""Process and fixture support for packaged CLI discovery execution verification."""

import json
import subprocess
import time
import uuid
from pathlib import Path
from typing import Callable, Final, Optional

PublicFixtureSecrets = dict[str, dict[str, str]]
PUBLIC_FIXTURE_SECRET_VALUES: Final[PublicFixtureSecrets] = {
    "PACKAGE_SECURITY_INSPECTION": {"source-open-password": "GridGrind-2026"},
}


class ArtifactRunner:
    """Runs one packaged CLI surface without exposing fixture credentials in progress output."""

    def __init__(
        self,
        mode: str,
        artifact_target: str,
        temp_root: Path,
        docker_run_user: str,
        heartbeat_seconds: int,
        progress: Callable[[str], None],
    ) -> None:
        self._mode = mode
        self._artifact_target = artifact_target
        self._temp_root = temp_root
        self._docker_run_user = docker_run_user
        self._heartbeat_seconds = heartbeat_seconds
        self._progress = progress

    def command(self, arguments: list[str], working_directory: Path) -> list[str]:
        if self._mode == "binary":
            return [self._artifact_target, *arguments]
        if self._mode == "jar":
            return ["java", "-jar", self._artifact_target, *arguments]
        if self._mode == "docker-image":
            docker_command = ["docker", "run", "--rm"]
            if self._docker_run_user:
                docker_command.extend(["--user", self._docker_run_user])
            return [
                *docker_command,
                "-v",
                f"{working_directory}:/work",
                self._artifact_target,
                *arguments,
            ]
        raise ValueError(f"unsupported launcher mode {self._mode}")

    def run(
        self,
        arguments: list[str],
        working_directory: Path,
        progress_label: Optional[str] = None,
    ) -> subprocess.CompletedProcess[str]:
        logs_directory = self._temp_root / "_subprocess_logs"
        logs_directory.mkdir(parents=True, exist_ok=True)
        stdout_path = logs_directory / f"{uuid.uuid4()}-stdout.log"
        stderr_path = logs_directory / f"{uuid.uuid4()}-stderr.log"
        command = self.command(arguments, working_directory)
        started_at = time.monotonic()
        next_heartbeat_at = self._heartbeat_seconds
        with stdout_path.open("w", encoding="utf-8") as stdout_handle, stderr_path.open(
            "w", encoding="utf-8"
        ) as stderr_handle:
            process = subprocess.Popen(
                command,
                cwd=working_directory,
                text=True,
                stdout=stdout_handle,
                stderr=stderr_handle,
            )
            while process.poll() is None:
                elapsed_seconds = time.monotonic() - started_at
                if progress_label is not None and elapsed_seconds >= next_heartbeat_at:
                    self._progress(
                        f"{progress_label} (still running after {int(elapsed_seconds)}s)"
                    )
                    next_heartbeat_at += self._heartbeat_seconds
                time.sleep(1)
        stdout = stdout_path.read_text(encoding="utf-8")
        stderr = stderr_path.read_text(encoding="utf-8")
        stdout_path.unlink(missing_ok=True)
        stderr_path.unlink(missing_ok=True)
        return subprocess.CompletedProcess(
            args=command,
            returncode=process.returncode,
            stdout=stdout,
            stderr=stderr,
        )

    def run_json(
        self,
        arguments: list[str],
        working_directory: Path,
        progress_label: Optional[str] = None,
    ) -> object:
        completed = self.run(arguments, working_directory, progress_label)
        if completed.returncode != 0:
            raise RuntimeError(
                f"command failed ({completed.returncode}): "
                + " ".join(self.command(arguments, working_directory))
                + f"\nstdout: {completed.stdout}\nstderr: {completed.stderr}"
            )
        try:
            return json.loads(completed.stdout)
        except json.JSONDecodeError as exception:
            raise RuntimeError(
                "command did not emit JSON: "
                + " ".join(self.command(arguments, working_directory))
                + f"\n{exception}\n{completed.stdout}"
            ) from exception

    def artifact_path(self, path: Path, workspace: Path) -> str:
        if self._mode in {"binary", "jar"}:
            return str(path)
        if self._mode == "docker-image":
            return str(path.relative_to(workspace))
        raise ValueError(f"unsupported launcher mode {self._mode}")

    def fixture_secret_provider(self, stable_id: str, workspace: Path) -> Optional[Path]:
        secret_values = PUBLIC_FIXTURE_SECRET_VALUES.get(stable_id)
        if secret_values is None:
            return None
        provider = workspace / "fixture-secrets"
        provider.mkdir()
        for reference, value in secret_values.items():
            (provider / reference).write_text(value, encoding="utf-8")
        return provider
