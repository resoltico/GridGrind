#!/usr/bin/env python3
"""Verify the actual managed JDK against the shared binary/source build tuple."""

from pathlib import Path
import re
import shlex
import shutil
import subprocess
import sys


def properties(text):
    return dict(
        line.partition("=")[::2]
        for line in text.splitlines()
        if line and not line.startswith("#") and "=" in line
    )


def verify_identity(pins, release, runtime):
    needed = {"TEMURIN_VERSION", "TEMURIN_BUILD_COMMIT", "TEMURIN_SOURCE_COMMIT"}
    missing = needed - pins.keys()
    if missing:
        raise ValueError(f"Missing managed JDK pins: {', '.join(sorted(missing))}")
    expected = pins["TEMURIN_VERSION"]
    required = {
        "IMPLEMENTOR": "Eclipse Adoptium",
        "IMPLEMENTOR_VERSION": f"Temurin-{expected}",
        "JAVA_RUNTIME_VERSION": expected,
        "BUILD_SOURCE": f'git:{pins["TEMURIN_BUILD_COMMIT"]}',
    }
    for key, value in required.items():
        if release.get(key) != value:
            raise ValueError(f"Managed JDK {key} must be {value}; found {release.get(key)!r}")
    source = release.get("SOURCE", "").rsplit(":", 1)[-1]
    if not re.fullmatch(r"[0-9a-f]{12,40}", source) or not pins["TEMURIN_SOURCE_COMMIT"].startswith(source):
        raise ValueError("Managed JDK SOURCE does not match the pinned full source commit")
    if runtime.get("java.runtime.version") != expected:
        raise ValueError(f"Running JDK must report java.runtime.version={expected}")
    if runtime.get("java.vendor") != "Eclipse Adoptium":
        raise ValueError("Running JDK must report java.vendor=Eclipse Adoptium")


def main():
    if len(sys.argv) != 2:
        raise ValueError("Usage: verify-managed-jdk.py JDK_HOME")
    home = Path(sys.argv[1]).resolve(strict=True)
    java = home / "bin/java"
    selected = shutil.which("java")
    if selected is None or Path(selected).resolve() != java.resolve():
        raise ValueError("PATH must select the verified JAVA_HOME/bin/java")
    pins_path = Path(__file__).resolve().parents[1] / "release/container-runtime.properties"
    pins = properties(pins_path.read_text(encoding="utf-8"))
    release = {key: shlex.split(value)[0] for key, value in
               properties((home / "release").read_text(encoding="utf-8")).items()}
    result = subprocess.run(
        [str(java), "-Duser.language=en", "-Duser.country=US", "-XshowSettings:properties", "-version"],
        check=True, capture_output=True, text=True, encoding="utf-8",
    )
    runtime = dict(re.findall(r"^\s*([\w.]+) = (.*)$", result.stderr, re.MULTILINE))
    verify_identity(pins, release, runtime)
    print(f'Verified Eclipse Temurin {pins["TEMURIN_VERSION"]}')


if __name__ == "__main__":
    try:
        main()
    except (OSError, ValueError, subprocess.CalledProcessError) as error:
        print(f"error: {error}", file=sys.stderr)
        sys.exit(1)
