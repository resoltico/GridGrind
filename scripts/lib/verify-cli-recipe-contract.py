#!/usr/bin/env python3
import json
import sys
from pathlib import Path


def load_json(index: int):
    return json.loads(Path(sys.argv[index]).read_text())


def die(message: str) -> None:
    print(f"error: {message}", file=sys.stderr)
    raise SystemExit(1)


def normalized_text(value: str) -> str:
    return " ".join(value.split())


recipe_catalog = load_json(1)
example_recipe_catalog_detail = load_json(2)
task_recipe_catalog_detail = load_json(3)
guidance_help_output = Path(sys.argv[4]).read_text()
source_types_group = load_json(5)
persistence_types_group = load_json(6)
step_types_group = load_json(7)
mutation_action_types_group = load_json(8)
assertion_types_group = load_json(9)
inspection_query_types_group = load_json(10)

normalized_guidance_help_output = normalized_text(guidance_help_output)
recipes = recipe_catalog["recipes"]
shipped_examples = [recipe for recipe in recipes if recipe.get("view") == "EXAMPLE"]
task_starters = [recipe for recipe in recipes if recipe.get("view") == "TASK_STARTER"]
expected_list_fields = {
    "view",
    "id",
    "requestFileName",
    "summary",
    "workspaceMode",
    "requiredWorkspacePaths",
}
catalog_groups = {
    source_types_group["group"]: {entry["id"] for entry in source_types_group["types"]},
    persistence_types_group["group"]: {entry["id"] for entry in persistence_types_group["types"]},
    step_types_group["group"]: {entry["id"] for entry in step_types_group["types"]},
    mutation_action_types_group["group"]: {
        entry["id"] for entry in mutation_action_types_group["types"]
    },
    assertion_types_group["group"]: {entry["id"] for entry in assertion_types_group["types"]},
    inspection_query_types_group["group"]: {
        entry["id"] for entry in inspection_query_types_group["types"]
    },
}

if recipe_catalog.get("protocolVersion") != "V2":
    die("recipe catalog no longer emits protocolVersion=V2")
if not recipes:
    die("recipe catalog recipes is empty")
if not shipped_examples:
    die("recipe catalog example view is empty")
if not task_starters:
    die("recipe catalog task-starter view is empty")

for recipe in shipped_examples + task_starters:
    recipe_id = recipe.get("id")
    if not recipe_id:
        die("recipe catalog contains a recipe with no id")
    view = recipe.get("view")
    request_file_name = recipe.get("requestFileName")
    summary = recipe.get("summary")
    guidance_snippet = normalized_text(
        f"- {recipe_id} ({view}) requestFileName: {request_file_name} "
        + f"workspace: {recipe.get('workspaceMode')} summary: {summary}"
    )
    if guidance_snippet not in normalized_guidance_help_output:
        die(f"guidance help no longer lists the recipe line for {recipe_id}")
    if set(recipe.keys()) != expected_list_fields:
        die(
            f"recipe catalog {view.lower()} entry {recipe_id} exposes unexpected public fields: "
            + f"{sorted(recipe.keys())}"
        )
    if recipe.get("workspaceMode") == "SELF_CONTAINED" and recipe.get(
        "requiredWorkspacePaths"
    ):
        die(f"self-contained recipe {recipe_id} must not publish requiredWorkspacePaths")
    if recipe.get("workspaceMode") == "REQUIRES_EXAMPLE_ASSETS" and not recipe.get(
        "requiredWorkspacePaths"
    ):
        die(f"asset-backed recipe {recipe_id} must publish requiredWorkspacePaths")

expected_required_paths = {
    "CUSTOM_XML": [
        "custom-xml-assets/custom-xml-mapping.xlsx",
        "custom-xml-assets/custom-xml-update.xml",
    ],
    "SOURCE_BACKED_INPUT": [
        "source-backed-input-assets/title.txt",
        "source-backed-input-assets/total-formula.txt",
        "source-backed-input-assets/payload.bin",
    ],
    "PACKAGE_SECURITY_INSPECTION": [
        "package-security-assets/gridgrind-package-security.xlsx",
    ],
}
for example_id, required_paths in expected_required_paths.items():
    entry = next(example for example in shipped_examples if example["id"] == example_id)
    if entry["requiredWorkspacePaths"] != required_paths:
        die(
            f"recipe catalog requiredWorkspacePaths drifted for {example_id}: "
            + f"{entry['requiredWorkspacePaths']}"
        )
    required_paths_snippet = normalized_text(
        f"requiredWorkspacePaths: {', '.join(required_paths)}"
    )
    if required_paths_snippet not in normalized_guidance_help_output:
        die(f"guidance help no longer lists requiredWorkspacePaths for {example_id}")

if example_recipe_catalog_detail.get("protocolVersion") != "V2":
    die("recipe catalog example lookup no longer emits protocolVersion=V2")
if example_recipe_catalog_detail.get("id") != "BUDGET":
    die("recipe catalog example lookup no longer returns the requested BUDGET detail")
if example_recipe_catalog_detail.get("view") != "EXAMPLE":
    die("recipe catalog example lookup drifted away from view=EXAMPLE")
if not example_recipe_catalog_detail.get("intentTags"):
    die("recipe catalog example lookup no longer publishes top-level intentTags")
if "requestProfile" not in example_recipe_catalog_detail:
    die("recipe catalog example lookup no longer publishes requestProfile")

if task_recipe_catalog_detail.get("protocolVersion") != "V2":
    die("recipe catalog task-starter lookup no longer emits protocolVersion=V2")
if task_recipe_catalog_detail.get("id") != "TABULAR_REPORT":
    die(
        "recipe catalog task-starter lookup no longer returns the requested "
        + "TABULAR_REPORT detail"
    )
if task_recipe_catalog_detail.get("view") != "TASK_STARTER":
    die("recipe catalog task-starter lookup drifted away from view=TASK_STARTER")
if not task_recipe_catalog_detail.get("intentTags"):
    die("recipe catalog task-starter lookup no longer publishes top-level intentTags")
if "discovery" in task_recipe_catalog_detail:
    die(
        "recipe catalog task-starter lookup still exposes the removed discovery wrapper"
    )
task_discovery = task_recipe_catalog_detail.get("discoveryProfile", {})
if not task_discovery.get("discoveryTerms"):
    die(
        "recipe catalog task-starter lookup no longer publishes "
        + "discoveryProfile.discoveryTerms"
    )
if "intentProfile" not in task_discovery:
    die(
        "recipe catalog task-starter lookup no longer publishes "
        + "discoveryProfile.intentProfile"
    )
if "requestProfile" not in task_recipe_catalog_detail:
    die("recipe catalog task-starter lookup no longer publishes requestProfile")

workflow = task_recipe_catalog_detail.get("workflow", {})
phases = workflow.get("phases", [])
if not phases:
    die("recipe catalog task-starter lookup contains no workflow phases")
for phase in phases:
    capability_refs = phase.get("capabilityRefs", [])
    if not capability_refs:
        die(
            "recipe catalog task-starter lookup phase "
            + f"{phase.get('label')} contains no capabilityRefs"
        )
    for capability_ref in capability_refs:
        group = capability_ref.get("group")
        capability_id = capability_ref.get("id")
        if group not in catalog_groups:
            die(
                "recipe catalog task-starter lookup references unknown protocol group "
                + f"{group}"
            )
        if capability_id not in catalog_groups[group]:
            die(
                "recipe catalog task-starter lookup references unknown protocol capability "
                + f"{group}:{capability_id}"
            )
