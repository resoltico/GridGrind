#!/usr/bin/env bash
# Guard high-value documentation contracts that must stay aligned with the executable surface.

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

export GRIDGRIND_REPO_ROOT="${repo_root}"

python3 - <<'PY'
from pathlib import Path
import json
import os
import re

root = Path(os.environ["GRIDGRIND_REPO_ROOT"])
frontmatter_files = sorted([*root.glob("docs/*.md"), root / "jazzer/README.md"])
live_contract_files = [
    root / "README.md",
    *(doc for doc in frontmatter_files if doc.name != "CHANGELOG_ARCHIVE.md"),
]
request_required = {"protocolVersion", "source", "persistence", "steps"}
current_protocol_version = "V3"
forbidden_secret_fields = {
    "password",
    "workbookPassword",
    "revisionsPassword",
    "keystorePassword",
    "keyPassword",
}


def contains_forbidden_secret_field(value):
    if isinstance(value, dict):
        if forbidden_secret_fields.intersection(value):
            return True
        return any(contains_forbidden_secret_field(nested) for nested in value.values())
    if isinstance(value, list):
        return any(contains_forbidden_secret_field(item) for item in value)
    return False

for doc in live_contract_files:
    for line in doc.read_text(encoding="utf-8").splitlines():
        if (
            '{"source":{"type":"NEW"},"steps":[]}' in line
            and "--doctor-request" in line
        ):
            raise SystemExit(
                "documentation still contains the obsolete no-envelope doctor-request example: "
                f"{doc.relative_to(root)}"
            )

for doc in frontmatter_files:
    text = doc.read_text(encoding="utf-8")
    if not text.startswith("---\n"):
        raise SystemExit(f"{doc.relative_to(root)} must start with AFAD frontmatter")
    if text.find("\n---\n", 4) == -1:
        raise SystemExit(f"{doc.relative_to(root)} has unterminated AFAD frontmatter")

for doc in live_contract_files:
    text = doc.read_text(encoding="utf-8")
    index = 0
    while True:
        start = text.find("```json", index)
        if start == -1:
            break
        end = text.find("```", start + 7)
        if end == -1:
            raise SystemExit(f"{doc.relative_to(root)} has an unterminated json fence")
        block = text[start + 7:end].strip()
        index = end + 3
        try:
            payload = json.loads(block)
        except Exception:
            continue
        if contains_forbidden_secret_field(payload):
            raise SystemExit(
                f"{doc.relative_to(root)} publishes an inline secret field; use a typed secret reference"
            )
        if isinstance(payload, dict) and "source" in payload and "steps" in payload:
            missing = sorted(request_required.difference(payload))
            if missing:
                raise SystemExit(
                    f"{doc.relative_to(root)} publishes a request-shaped json block missing {missing}"
                )
            if payload["protocolVersion"] != current_protocol_version:
                raise SystemExit(
                    f"{doc.relative_to(root)} publishes protocolVersion="
                    f"{payload['protocolVersion']!r}; expected {current_protocol_version!r}"
                )
            encryption = (
                payload.get("persistence", {})
                .get("security", {})
                .get("encryption")
            )
            if isinstance(encryption, dict) and encryption.get("type") == "ENCRYPT":
                settings = encryption.get("encryption")
                if not isinstance(settings, dict):
                    raise SystemExit(
                        f"{doc.relative_to(root)} publishes ENCRYPT without its encryption settings"
                    )
                encryption_missing = sorted({"passwordRef"}.difference(settings))
                if encryption_missing:
                    raise SystemExit(
                        f"{doc.relative_to(root)} publishes persistence encryption without {encryption_missing}"
                    )

coverage_document = (root / "docs/DEVELOPER_JAZZER_COVERAGE.md").read_text(encoding="utf-8")
fuzz_input_directories = {
    "protocol-request": root / "jazzer/src/fuzz/resources/dev/erst/gridgrind/jazzer/protocol/ProtocolRequestFuzzTestInputs",
    "protocol-workflow": root / "jazzer/src/fuzz/resources/dev/erst/gridgrind/jazzer/protocol/OperationWorkflowFuzzTestInputs",
    "engine-command-sequence": root / "jazzer/src/fuzz/resources/dev/erst/gridgrind/jazzer/engine/WorkbookCommandSequenceFuzzTestInputs",
    "xlsx-roundtrip": root / "jazzer/src/fuzz/resources/dev/erst/gridgrind/jazzer/engine/XlsxRoundTripFuzzTestInputs",
}
total_promoted_inputs = 0
for target, input_directory in fuzz_input_directories.items():
    actual_inputs = {path.name for path in input_directory.rglob("*") if path.is_file()}
    total_promoted_inputs += len(actual_inputs)
    summary_line = next(
        (line for line in coverage_document.splitlines() if line.startswith(f"| `{target}` ")),
        None,
    )
    if summary_line is None:
        raise SystemExit(f"Jazzer coverage inventory is missing its {target} summary row")
    if int(summary_line.rsplit("|", 2)[1].strip()) != len(actual_inputs):
        raise SystemExit(f"Jazzer coverage inventory has a stale {target} summary count")
    section = re.search(
        rf"^### `{re.escape(target)}` \(\d+\)$(.*?)(?=^### |\Z)",
        coverage_document,
        re.MULTILINE | re.DOTALL,
    )
    if section is None:
        raise SystemExit(f"Jazzer coverage inventory is missing its {target} input section")
    documented_inputs = set(re.findall(r"^- `([^`]+)`$", section.group(1), re.MULTILINE))
    if documented_inputs != actual_inputs:
        raise SystemExit(f"Jazzer coverage inventory has a stale {target} input list")

if f"{total_promoted_inputs} total across harnesses" not in coverage_document:
    raise SystemExit("Jazzer coverage inventory has a stale total promoted-input count")
PY

printf 'documentation contract regression: success\n'
