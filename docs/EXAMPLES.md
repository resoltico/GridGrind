---
afad: "5.0.1"
version: "0.73.0"
domain: EXAMPLES
updated: "2026-08-27"
route:
  keywords: [gridgrind, examples, print-recipe, request fixtures, package security, java authoring]
  questions: ["what examples ship with gridgrind", "what is the difference between built-in and checked-in examples", "how do i run the java example", "how do i refresh the example fixtures"]
---

# Example Guide

**Purpose**: Map the shipped example surfaces, explain how their paths resolve, and show how to
refresh and verify them.
**Fastest artifact-native path**: `gridgrind --print-recipe --lookup <ID> --response request.json`
**Docker `:latest` note**: for first-contact artifact runs, prefer
`docker run --pull=always --rm ghcr.io/resoltico/gridgrind:latest ...` or refresh once with
`docker pull ghcr.io/resoltico/gridgrind:latest` before using plain `docker run ...:latest`
**Checkout-owned fixtures**: [`../examples/`](../examples/)

GridGrind ships the same example workflows in two forms:

- **Built-in artifact examples** from `gridgrind --print-recipe --lookup <ID> --response request.json`.
  These are designed to run from an artifact working directory and use request-relative paths such
  as `generated-workbooks/...`. The machine-readable catalog publishes one portable
  `requestFileName`, while any repo-backed asset requirements are published separately through
  `requiredWorkspacePaths`. When one example uses the default execution path and empty evaluator
  environment, the printed JSON omits those two top-level blocks and keeps only the minimal
  request envelope.
  They are not all equally portable: most are self-contained in a blank working directory, while a
  few are intentionally repo-asset-backed.
- **Checked-in repository fixtures** under [`../examples/`](../examples/). These are generated from
  the same CLI recipe registry, but their relative paths are rooted from the request file's own
  directory so they run in place from a repository checkout.

Cell-reading examples intentionally stay compact unless they are demonstrating a richer surface:
`GET_CELLS`, `GET_WINDOW`, and `GET_SHEET_SCHEMA` omit `projection` for the default `[VALUE]`
readback, window examples stay sparse unless `includeBlanks` is set to `true`, and every shipped
cell-reading example stays comfortably inside the shared deterministic 250,000-cell read cap.

## Path Rules

- Built-in examples are for the packaged `gridgrind` launcher, the release JAR, the Docker image,
  or `:cli:run` when you first print the example into your own working directory.
- Self-contained built-ins need no workbook or asset inputs. Create their `generated-workbooks/`
  output directory before the first run, then execute them from an otherwise blank artifact workspace.
- Repo-asset-backed built-ins require the matching asset paths named in `requiredWorkspacePaths` to exist
  in the working directory before you run them.
- Any example that saves a workbook writes under `generated-workbooks/` beside the request file.
  Create that parent directory first; GridGrind intentionally does not create it through an
  unverified path lookup.
- Checked-in `examples/*.json` therefore persist into `examples/generated-workbooks/` because the
  request files themselves live under `examples/`.
- Asset-backed checked-in examples keep sibling assets beside the requests:
  - [`../examples/custom-xml-assets/`](../examples/custom-xml-assets/)
  - [`../examples/source-backed-input-assets/`](../examples/source-backed-input-assets/)
  - [`../examples/package-security-assets/`](../examples/package-security-assets/)

## Built-In Example Portability

Self-contained built-ins execute after `mkdir -p generated-workbooks` in an otherwise blank artifact workspace and `--print-recipe --lookup <ID>`:

| Built-in ID | Matching fixture | Notes |
|:------------|:-----------------|:------|
| `BUDGET` | [`../examples/budget-request.json`](../examples/budget-request.json) | self-contained budget walkthrough |
| `WORKBOOK_HEALTH` | [`../examples/workbook-health-request.json`](../examples/workbook-health-request.json) | no-save workbook-health flow |
| `SHEET_MAINTENANCE` | [`../examples/sheet-maintenance-request.json`](../examples/sheet-maintenance-request.json) | copy-sheet and maintenance flow |
| `ASSERTION` | [`../examples/assertion-request.json`](../examples/assertion-request.json) | mutate-then-assert walkthrough |
| `ARRAY_FORMULA` | [`../examples/array-formula-request.json`](../examples/array-formula-request.json) | array-group authoring and readback |
| `SIGNATURE_LINE` | [`../examples/signature-line-request.json`](../examples/signature-line-request.json) | drawing and signature metadata |
| `LARGE_FILE_MODES` | [`../examples/large-file-modes-request.json`](../examples/large-file-modes-request.json) | `STREAMING_WRITE` plus summary readback |
| `CHART` | [`../examples/chart-request.json`](../examples/chart-request.json) | supported chart authoring |
| `PIVOT` | [`../examples/pivot-request.json`](../examples/pivot-request.json) | pivot authoring and health analysis |
| `INTROSPECTION_ANALYSIS` | [`../examples/introspection-analysis-request.json`](../examples/introspection-analysis-request.json) | inspection-heavy analysis surface |

Repo-asset-backed built-ins still use `--print-recipe --lookup <ID>`, but they also require the
copied asset paths named in `requiredWorkspacePaths`:

| Built-in ID | Matching fixture | Required assets |
|:------------|:-----------------|:----------------|
| `CUSTOM_XML` | [`../examples/custom-xml-request.json`](../examples/custom-xml-request.json) | [`../examples/custom-xml-assets/`](../examples/custom-xml-assets/) |
| `FILE_HYPERLINK_HEALTH` | [`../examples/file-hyperlink-health-request.json`](../examples/file-hyperlink-health-request.json) | [`../examples/file-hyperlink-assets/request-label.txt`](../examples/file-hyperlink-assets/request-label.txt) |
| `SOURCE_BACKED_INPUT` | [`../examples/source-backed-input-request.json`](../examples/source-backed-input-request.json) | [`../examples/source-backed-input-assets/`](../examples/source-backed-input-assets/) |
| `PACKAGE_SECURITY_INSPECTION` | [`../examples/package-security-inspect-request.json`](../examples/package-security-inspect-request.json) | [`../examples/package-security-assets/`](../examples/package-security-assets/) |

The CLI help now prints each built-in example with its `advisory`, and asset-backed entries
also print their exact `requiredWorkspacePaths`, so artifact-only workspaces do not silently assume every
example is self-contained.

The machine-readable CLI recipe catalog exposes stable example ids, file names, summaries, a
portable `requestFileName` plus `advisory` contract, and exact
`requiredWorkspacePaths` for asset-backed examples.
`SELF_CONTAINED` means the printed request runs from a blank artifact workspace;
`REQUIRES_EXAMPLE_ASSETS` means the request expects copied `examples/` assets beside the request
file, and `requiredWorkspacePaths` names those files directly.
Print it directly with `gridgrind --print-recipe-catalog --response recipes.json`.

## JSON Request Fixtures

| Example | Shape | Notes |
|:--------|:------|:------|
| [`../examples/budget-request.json`](../examples/budget-request.json) | create and save | budget walkthrough; matches built-in `BUDGET` |
| [`../examples/assertion-request.json`](../examples/assertion-request.json) | no-save verify | ordered mutate-then-assert flow |
| [`../examples/workbook-health-request.json`](../examples/workbook-health-request.json) | no-save inspect | compact workbook-health workflow |
| [`../examples/sheet-maintenance-request.json`](../examples/sheet-maintenance-request.json) | create and save | copy-sheet and workbook-maintenance flow |
| [`../examples/array-formula-request.json`](../examples/array-formula-request.json) | no-save inspect | array-formula authoring and group readback |
| [`../examples/chart-request.json`](../examples/chart-request.json) | create and save | chart authoring with factual readback |
| [`../examples/pivot-request.json`](../examples/pivot-request.json) | create and save | pivot authoring plus pivot-health analysis |
| [`../examples/file-hyperlink-health-request.json`](../examples/file-hyperlink-health-request.json) | create and save | request-asset text plus workbook-relative hyperlink analysis |
| [`../examples/introspection-analysis-request.json`](../examples/introspection-analysis-request.json) | create and save | inspection-heavy workbook analysis surface |
| [`../examples/large-file-modes-request.json`](../examples/large-file-modes-request.json) | create and save | `STREAMING_WRITE` and recalculation-open flagging |
| [`../examples/source-backed-input-request.json`](../examples/source-backed-input-request.json) | no-save inspect | sibling file-backed text, formula, and binary payloads |
| [`../examples/custom-xml-request.json`](../examples/custom-xml-request.json) | existing workbook | sibling custom-XML assets and XML import/export |
| [`../examples/signature-line-request.json`](../examples/signature-line-request.json) | create and save | signature-line and drawing-anchor surface |
| [`../examples/package-security-inspect-request.json`](../examples/package-security-inspect-request.json) | existing encrypted workbook | reopens the committed encrypted asset under `package-security-assets/` |

## Java Authoring Example

The Java-first example lives at [../examples/java-authoring-workflow.java](../examples/java-authoring-workflow.java).
It is compiled and executed by `:authoring-java:test`, not merely syntax-checked.

If you want to run that example against the repository checkout directly, pass the `examples/`
directory as the workspace root so the example can read the committed
[../examples/authored-inputs/item.txt](../examples/authored-inputs/item.txt) file.

## Refresh And Verification

Refresh the checkout-rooted request fixtures and the generated package-security workbook asset with:

```bash
./scripts/sync-generated-examples.sh
```

Treat the checked-in `examples/*.json` fixtures and the generated package-security workbook as derived artifacts from that script and the CLI recipe registry's example view; do not hand-edit them in place.

The authoritative verification loop for the shipped examples is:

```bash
./gradlew :cli:test --tests dev.erst.gridgrind.cli.discovery.ExampleRequestFixturesTest
./gradlew :executor:test --tests dev.erst.gridgrind.engine.runtime.ExampleExecutionFixturesTest
./gradlew :authoring-java:test --tests dev.erst.gridgrind.authoring.GridGrindPlanTest
./scripts/verify-cli-discovery-execution.sh ./cli/build/libs/gridgrind.jar
```

For a direct packaged-CLI spot check from a repository checkout, this also works:

```bash
java -jar cli/build/libs/gridgrind.jar \
  --request examples/budget-request.json \
  --response tmp/example-budget-response.json
```
