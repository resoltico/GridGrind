---
afad: "3.5"
version: "0.61.0"
domain: EXAMPLES
updated: "2026-04-25"
route:
  keywords: [gridgrind, examples, print-example, request fixtures, package security, java authoring]
  questions: ["what examples ship with gridgrind", "what is the difference between built-in and checked-in examples", "how do i run the java example", "how do i refresh the example fixtures"]
---

# Example Guide

**Purpose**: Map the shipped example surfaces, explain how their paths resolve, and show how to
refresh and verify them.
**Fastest artifact-native path**: `gridgrind --print-example <ID> --response request.json`
**Docker `:latest` note**: for first-contact artifact runs, prefer
`docker run --pull=always --rm ghcr.io/resoltico/gridgrind:latest ...` or refresh once with
`docker pull ghcr.io/resoltico/gridgrind:latest` before using plain `docker run ...:latest`
**Checkout-owned fixtures**: [`../examples/`](../examples/)

GridGrind ships the same example workflows in two forms:

- **Built-in artifact examples** from `gridgrind --print-example <ID> --response request.json`.
  These are designed to run from an artifact working directory and use artifact-rooted paths such
  as `examples/...` or `cli/build/generated-workbooks/...`. They are not all equally portable:
  most are self-contained in a blank working directory, while a few are intentionally
  repo-asset-backed.
- **Checked-in repository fixtures** under [`../examples/`](../examples/). These are generated from
  the same contract-owned registry, but their relative paths are rooted from the request file's own
  directory so they run in place from a repository checkout.

## Path Rules

- Built-in examples are for the release JAR, Docker image, or `:cli:run` when you first print the
  example into your own working directory.
- Self-contained built-ins can run from a blank artifact workspace after you print the request.
- Repo-asset-backed built-ins require the matching `examples/` asset directories to exist in the
  working directory before you run them.
- Checked-in `examples/*.json` are checkout-rooted fixtures. Persisted outputs intentionally use
  `../cli/build/generated-workbooks/...` because the request files themselves live under
  `examples/`.
- Asset-backed checked-in examples keep sibling assets beside the requests:
  - [`../examples/custom-xml-assets/`](../examples/custom-xml-assets/)
  - [`../examples/source-backed-input-assets/`](../examples/source-backed-input-assets/)
  - [`../examples/package-security-assets/`](../examples/package-security-assets/)

## Built-In Example Portability

Self-contained built-ins execute from a blank artifact workspace after `--print-example`:

| Built-in ID | Matching fixture | Notes |
|:------------|:-----------------|:------|
| `BUDGET` | [`../examples/budget-request.json`](../examples/budget-request.json) | first-run starter |
| `WORKBOOK_HEALTH` | [`../examples/workbook-health-request.json`](../examples/workbook-health-request.json) | no-save workbook-health flow |
| `SHEET_MAINTENANCE` | [`../examples/sheet-maintenance-request.json`](../examples/sheet-maintenance-request.json) | copy-sheet and maintenance flow |
| `ASSERTION` | [`../examples/assertion-request.json`](../examples/assertion-request.json) | mutate-then-assert walkthrough |
| `ARRAY_FORMULA` | [`../examples/array-formula-request.json`](../examples/array-formula-request.json) | array-group authoring and readback |
| `SIGNATURE_LINE` | [`../examples/signature-line-request.json`](../examples/signature-line-request.json) | drawing and signature metadata |
| `LARGE_FILE_MODES` | [`../examples/large-file-modes-request.json`](../examples/large-file-modes-request.json) | `STREAMING_WRITE` plus summary readback |
| `CHART` | [`../examples/chart-request.json`](../examples/chart-request.json) | supported chart authoring |
| `PIVOT` | [`../examples/pivot-request.json`](../examples/pivot-request.json) | pivot authoring and health analysis |
| `FILE_HYPERLINK_HEALTH` | [`../examples/file-hyperlink-health-request.json`](../examples/file-hyperlink-health-request.json) | file/document hyperlink analysis |
| `INTROSPECTION_ANALYSIS` | [`../examples/introspection-analysis-request.json`](../examples/introspection-analysis-request.json) | inspection-heavy analysis surface |

Repo-asset-backed built-ins still use `--print-example <ID>`, but they also require copied
`examples/` assets in the working directory:

| Built-in ID | Matching fixture | Required assets |
|:------------|:-----------------|:----------------|
| `CUSTOM_XML` | [`../examples/custom-xml-request.json`](../examples/custom-xml-request.json) | [`../examples/custom-xml-assets/`](../examples/custom-xml-assets/) |
| `SOURCE_BACKED_INPUT` | [`../examples/source-backed-input-request.json`](../examples/source-backed-input-request.json) | [`../examples/source-backed-input-assets/`](../examples/source-backed-input-assets/) |
| `PACKAGE_SECURITY_INSPECTION` | [`../examples/package-security-inspect-request.json`](../examples/package-security-inspect-request.json) | [`../examples/package-security-assets/`](../examples/package-security-assets/) |

The CLI help summaries already label those three built-ins as repo-asset-backed so artifact-only
workspaces do not silently assume every example is self-contained.

The machine-readable protocol catalog exposes that same portability split through
`shippedExamples[*].workspaceMode` and `shippedExamples[*].requiredPaths`, so agents do not have to
infer prerequisites from prose alone.

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
| [`../examples/file-hyperlink-health-request.json`](../examples/file-hyperlink-health-request.json) | create and save | hyperlink authoring and hyperlink-health analysis |
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

The authoritative verification loop for the shipped examples is:

```bash
./gradlew :contract:test --tests dev.erst.gridgrind.contract.json.ExampleRequestFixturesTest
./gradlew :executor:test --tests dev.erst.gridgrind.executor.ExampleExecutionFixturesTest
./gradlew :authoring-java:test --tests dev.erst.gridgrind.authoring.GridGrindPlanTest
```

For a direct packaged-CLI spot check from a repository checkout, this also works:

```bash
java -jar cli/build/libs/gridgrind.jar \
  --request examples/budget-request.json \
  --response tmp/example-budget-response.json
```
