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

root = Path(os.environ["GRIDGRIND_REPO_ROOT"])
frontmatter_files = sorted([*root.glob("docs/*.md"), root / "jazzer/README.md"])
request_required = {
    "protocolVersion",
    "source",
    "persistence",
    "execution",
    "formulaEnvironment",
    "steps",
}

for doc in [root / "README.md", *frontmatter_files]:
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

docs_with_json = [root / "README.md", *frontmatter_files]
for doc in docs_with_json:
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
        if isinstance(payload, dict) and "source" in payload and "steps" in payload:
            missing = sorted(request_required.difference(payload))
            if missing:
                raise SystemExit(
                    f"{doc.relative_to(root)} publishes a request-shaped json block missing {missing}"
                )
            encryption = (
                payload.get("persistence", {})
                .get("security", {})
                .get("encryption")
            )
            if isinstance(encryption, dict):
                encryption_missing = sorted({"password", "mode"}.difference(encryption))
                if encryption_missing:
                    raise SystemExit(
                        f"{doc.relative_to(root)} publishes persistence encryption without {encryption_missing}"
                    )
PY

printf 'documentation contract regression: success\n'
