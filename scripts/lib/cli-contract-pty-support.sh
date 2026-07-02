#!/usr/bin/env bash
# PTY-backed interactive no-arg verification shared by the packaged CLI contract gate.

verify_interactive_noarg_failure() {
    local expected_failure_path=$1
    shift

    "${cli_contract_python}" - "${expected_failure_path}" "$@" <<'PY'
import errno
import json
import os
import pty
import select
import subprocess
import sys
import time
from pathlib import Path

expected_failure = Path(sys.argv[1]).read_text().replace('\r', '').rstrip('\n')
expected_failure_json = json.loads(expected_failure)
command = sys.argv[2:]
timeout_seconds = 10.0

def first_json_document(text: str):
    stripped = text.lstrip()
    if not stripped:
        raise ValueError('no JSON payload found')
    decoder = json.JSONDecoder()
    value, _ = decoder.raw_decode(stripped)
    return value

master_fd, slave_fd = pty.openpty()
process = None
captured = bytearray()

try:
    process = subprocess.Popen(
        command,
        stdin=slave_fd,
        stdout=slave_fd,
        stderr=slave_fd,
        close_fds=True,
    )
finally:
    os.close(slave_fd)

try:
    deadline = time.monotonic() + timeout_seconds
    while True:
        if process.poll() is not None:
            break
        if time.monotonic() >= deadline:
            process.terminate()
            try:
                process.wait(timeout=2.0)
            except subprocess.TimeoutExpired:
                process.kill()
                process.wait(timeout=2.0)
            output = captured.decode('utf-8', 'replace').replace('\r', '').rstrip('\n')
            print(
                'error: interactive no-arg invocation did not exit promptly; '
                + f'partial output: {output[:400]!r}',
                file=sys.stderr,
            )
            raise SystemExit(1)
        ready, _, _ = select.select([master_fd], [], [], 0.1)
        if not ready:
            continue
        try:
            chunk = os.read(master_fd, 4096)
        except OSError as exc:
            if exc.errno == errno.EIO:
                break
            raise
        if not chunk:
            break
        captured.extend(chunk)

    while True:
        try:
            chunk = os.read(master_fd, 4096)
        except OSError as exc:
            if exc.errno == errno.EIO:
                break
            raise
        if not chunk:
            break
        captured.extend(chunk)
finally:
    os.close(master_fd)

process.wait(timeout=2.0)
output = captured.decode('utf-8', 'replace').replace('\r', '').rstrip('\n')
if process.returncode != 2:
    print(
        'error: interactive no-arg invocation exited with the wrong status '
        + f'({process.returncode}) with output {output[:400]!r}',
        file=sys.stderr,
    )
    raise SystemExit(1)
lines = [line for line in output.split('\n') if line]
if not lines:
    print(
        'error: interactive no-arg invocation emitted no product output',
        file=sys.stderr,
    )
    raise SystemExit(1)
try:
    actual_failure_json = first_json_document(output)
except (json.JSONDecodeError, ValueError):
    print(
        'error: interactive no-arg invocation did not start with the expected JSON failure report',
        file=sys.stderr,
    )
    raise SystemExit(1)
if actual_failure_json != expected_failure_json:
    print(
        'error: interactive no-arg invocation output differed from the expected failure report',
        file=sys.stderr,
    )
    raise SystemExit(1)
PY
}
