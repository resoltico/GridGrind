#!/usr/bin/env python3
"""Write the explicit test-host grant required by one generated V3 request fixture."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Final, TypeAlias, cast

JsonArray: TypeAlias = list["JsonValue"]
JsonObject: TypeAlias = dict[str, "JsonValue"]
JsonValue: TypeAlias = JsonArray | JsonObject | str | int | float | bool | None
FILE_SOURCE_TYPES: Final = frozenset({"EXISTING", "UTF8_FILE"})
FILE_SOURCE_PARENT_KEYS: Final = frozenset({"source", "payload"})
OPERATION_KEYS: Final = frozenset({"action", "assertion", "query"})


def main() -> int:
    arguments = parse_arguments()
    request_path = arguments.request.resolve()
    request = read_request(request_path)
    grant = derive_grant(request, request_path.parent, arguments.grant_working_directory)
    if arguments.prepare_save_as_parent:
        prepare_save_as_parent(request, request_path.parent)
    arguments.output.parent.mkdir(parents=True, exist_ok=True)
    arguments.output.write_text(json.dumps(grant, indent=2) + "\n", encoding="utf-8")
    return 0


def parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="write the narrow explicit host grant for one V3 fixture request"
    )
    parser.add_argument("--request", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument(
        "--grant-working-directory",
        type=Path,
        help="render contained grant paths relative to this host working directory",
    )
    parser.add_argument(
        "--prepare-save-as-parent",
        action="store_true",
        help="create the request-relative parent directory for a SAVE_AS fixture output",
    )
    return parser.parse_args()


def read_request(request_path: Path) -> JsonObject:
    try:
        parsed = json.loads(request_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exception:
        raise SystemExit(f"could not read fixture request {request_path}: {exception}") from exception
    if not isinstance(parsed, dict):
        raise SystemExit(f"fixture request {request_path} must contain one JSON object")
    return cast(JsonObject, parsed)


def derive_grant(
    request: JsonObject,
    request_directory: Path,
    grant_working_directory: Path | None,
) -> JsonObject:
    resources: set[Path] = set()
    operation_ids: set[str] = set()
    secret_references: set[str] = set()
    collect_request_facts(request, request_directory, resources, operation_ids, secret_references)
    return {
        "readableResources": [
            {
                "type": "FILE",
                "path": render_grant_path(path, grant_working_directory),
            }
            for path in sorted(resources)
        ],
        "operationIds": sorted(operation_ids),
        "targetAuthority": {"type": "WORKBOOK_WIDE"},
        "publicationAuthority": publication_authority(
            request, request_directory, grant_working_directory
        ),
        "allowedSecretReferences": [
            {"id": reference} for reference in sorted(secret_references)
        ],
        "acceptancePolicy": {"type": "MINIMUM_ONLY"},
    }


def collect_request_facts(
    value: JsonValue,
    request_directory: Path,
    resources: set[Path],
    operation_ids: set[str],
    secret_references: set[str],
    parent_key: str | None = None,
) -> None:
    if isinstance(value, dict):
        value_type = value.get("type")
        if parent_key in OPERATION_KEYS and isinstance(value_type, str):
            operation_ids.add(value_type)
        if value_type in FILE_SOURCE_TYPES or (
            value_type == "FILE" and parent_key in FILE_SOURCE_PARENT_KEYS
        ):
            raw_path = value.get("path")
            if isinstance(raw_path, str) and raw_path:
                resources.add(resolve_request_path(raw_path, request_directory))
        for key, nested in value.items():
            if key.endswith("Ref") and isinstance(nested, str) and nested:
                secret_references.add(nested)
            if key.endswith("Ref") and isinstance(nested, dict):
                reference_id = nested.get("id")
                if isinstance(reference_id, str) and reference_id:
                    secret_references.add(reference_id)
            collect_request_facts(
                nested,
                request_directory,
                resources,
                operation_ids,
                secret_references,
                key,
            )
    elif isinstance(value, list):
        for nested in value:
            collect_request_facts(
                nested,
                request_directory,
                resources,
                operation_ids,
                secret_references,
                parent_key,
            )


def publication_authority(
    request: JsonObject,
    request_directory: Path,
    grant_working_directory: Path | None,
) -> JsonObject:
    persistence = request.get("persistence")
    if not isinstance(persistence, dict):
        raise SystemExit("fixture request persistence must be an object")
    persistence_type = persistence.get("type")
    if persistence_type == "NONE":
        return {"type": "NONE"}
    if persistence_type == "OVERWRITE":
        return {"type": "OVERWRITE_SOURCE"}
    if persistence_type != "SAVE_AS":
        raise SystemExit(f"fixture request has unknown persistence type {persistence_type!r}")
    raw_path = persistence.get("path")
    collision_policy = persistence.get("ifExists")
    if not isinstance(raw_path, str) or not raw_path:
        raise SystemExit("fixture SAVE_AS persistence requires a nonempty path")
    if not isinstance(collision_policy, str):
        raise SystemExit("fixture SAVE_AS persistence requires an ifExists value")
    return {
        "type": "SAVE_AS",
        "path": render_grant_path(
            resolve_request_path(raw_path, request_directory), grant_working_directory
        ),
        "ifExists": collision_policy,
    }


def prepare_save_as_parent(request: JsonObject, request_directory: Path) -> None:
    persistence = request.get("persistence")
    if not isinstance(persistence, dict) or persistence.get("type") != "SAVE_AS":
        return
    raw_path = persistence.get("path")
    if not isinstance(raw_path, str) or not raw_path:
        raise SystemExit("fixture SAVE_AS persistence requires a nonempty path")
    resolve_request_path(raw_path, request_directory).parent.mkdir(parents=True, exist_ok=True)


def resolve_request_path(raw_path: str, request_directory: Path) -> Path:
    candidate = Path(raw_path)
    return (candidate if candidate.is_absolute() else request_directory / candidate).resolve()


def render_grant_path(path: Path, grant_working_directory: Path | None) -> str:
    if grant_working_directory is None:
        return str(path)
    try:
        return str(path.relative_to(grant_working_directory.resolve()))
    except ValueError as exception:
        raise SystemExit(
            f"fixture path {path} escapes grant working directory {grant_working_directory}"
        ) from exception


if __name__ == "__main__":
    sys.exit(main())
