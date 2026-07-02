---
afad: "4.0"
version: "0.71.0"
domain: QUICK_START
updated: "2026-06-29"
route:
  keywords: [gridgrind, quick start, first run, docker, jar, xlsx, example, response]
  questions: ["how do i do a first run with gridgrind", "what is the fastest way to try gridgrind", "how do i run the shipped examples", "how do i get my first successful gridgrind run"]
---

# Quick Start

Get to a first successful GridGrind run with the least setup and the least guesswork. The fastest path is to generate the built-in `BUDGET` example directly from the artifact: `--print-example --lookup BUDGET --response budget-request.json`. If you are already in a repo checkout, the matching JSON also lives at [../examples/budget-request.json](../examples/budget-request.json).
Generated example JSON already includes the top-level request envelope and omits the default
`execution` and `formulaEnvironment` blocks, so the first printed request is ready for copy-paste
edits without hand-authoring boilerplate defaults yourself.

---

## What You Need

- A GridGrind runtime: the packaged `gridgrind` launcher, the Docker image, or the release JAR
- A working directory where GridGrind can read the request file and write the response file
- One example request to start from: the built-in `BUDGET` example emitted by `--print-example --lookup BUDGET --response budget-request.json`, or [budget-request.json](../examples/budget-request.json) when you are already in a repo checkout

GridGrind supports `.xlsx` workbooks only.

If you are starting from the release artifact alone, generate the request into your working
directory first so the later `--request` path already exists: `gridgrind --print-example --lookup BUDGET --response budget-request.json`.
When you later run `--request budget-request.json`, relative paths inside that JSON request follow
the request file's directory. If you prefer to pipe request JSON on stdin, pass
`--execution-root <path>` and those same relative request-owned paths resolve from that explicit
directory. The separate CLI path flags `--response`, `--execution-root`, and `--temp-root` follow
the shell working directory.

## Pick One Run Path

### Docker

If Docker is the easiest path on your machine, the safest first-contact command is:

```bash
docker run --pull=always --rm ghcr.io/resoltico/gridgrind:latest --help
```

If you prefer to refresh once and then run several commands locally, pull the current `latest`
first:

```bash
docker pull ghcr.io/resoltico/gridgrind:latest
```

Docker does not automatically refresh a locally cached `latest` tag during a plain
`docker run ...:latest`. For first-contact copy-paste commands, keep `--pull=always` in place.
After an explicit `docker pull ghcr.io/resoltico/gridgrind:latest`, you can drop `--pull=always`
during repeated local runs if you want.

The published image already includes the font stack required for signature-line preview
generation, so signature-line requests work in Docker without extra image customization.
Mount the host working directory at `/work` and rely on the image's prepared `WORKDIR` so
relative CLI paths resolve inside that mounted directory without a separate `-w` override.

If you are already in a repository checkout and want the same runtime container without fetching a
release asset first, build the root Dockerfile directly:

```bash
docker buildx build --load -t gridgrind-local .
docker run --rm gridgrind-local --help
```

### Release JAR

If you want the standalone JAR, download it from the
[latest release](https://github.com/resoltico/GridGrind/releases/latest).

The JAR path requires Java 26.
`java -jar gridgrind.jar --help` prints the short synopsis. `--help-protocol` prints the
authoritative CLI/request contract, and `--help-guidance` prints workflows plus example guidance.
A bare `java -jar gridgrind.jar` invocation now expects a real request document on stdin or
`--request <path>` and exits non-zero when neither is present.

## First Successful Run

Use the built-in `BUDGET` example for the first pass. It writes a sample workbook and a JSON
response, so you can see both the output file and the run result. If you are already in a repo
checkout, [budget-request.json](../examples/budget-request.json) is the matching checked-in copy.
`BUDGET` is intentionally self-contained in a blank artifact workspace. A few other built-in
examples are repo-asset-backed and expect the copied asset paths named by
`requiredWorkspacePaths`; [EXAMPLES.md](./EXAMPLES.md) calls those out explicitly, and
`--print-example-catalog` exposes that distinction through each example's `requestFileName`,
`workspaceMode`, and asset-backed `requiredWorkspacePaths`.

### Docker Example

Generate the built-in request once, then run it from the current directory:

```bash
docker run --pull=always --rm ghcr.io/resoltico/gridgrind:latest --print-example --lookup BUDGET \
  --response budget-request.json

docker run --pull=always --rm -i \
  -v "$(pwd)":/work \
  ghcr.io/resoltico/gridgrind:latest \
  --request budget-request.json \
  --response response.json
```

### Docker Example From A Repository Checkout

Build the runtime image locally once, then use that local tag in the same mounted-directory flow:

```bash
docker buildx build --load -t gridgrind-local .

docker run --rm gridgrind-local --print-example --lookup BUDGET \
  --response budget-request.json

docker run --rm -i \
  -v "$(pwd)":/work \
  gridgrind-local \
  --request budget-request.json \
  --response response.json
```

### JAR Example

Replace `gridgrind.jar` with the downloaded JAR filename if it differs on your machine.

```bash
java -jar gridgrind.jar --print-example --lookup BUDGET --response budget-request.json

java -jar gridgrind.jar \
  --request budget-request.json \
  --response response.json
```

## What To Look For

After a successful run:

- `response.json` should report `status: "SUCCEEDED"`
- the workbook is saved to the path set in `persistence.path` inside the request JSON; open the generated `budget-request.json` and edit that field to control the output location
- if the run fails, GridGrind returns a structured error response instead of saving a partial workbook

## Good Second Steps

- Want the full example map, path rules, and refresh flow: [EXAMPLES.md](./EXAMPLES.md)
- Want GridGrind to explain itself from the artifact instead of from prose:
  - `--print-task-catalog --response tasks.json` lists the CLI-owned high-level office-work tasks, including dashboards, pivot reports, custom XML workflows, workbook maintenance, and drawing/signature flows.
  - `--print-task-plan --lookup DASHBOARD --response dashboard-request.json` emits one validated executable starter request for one task id.
  - `--print-task-keyword-match --query "monthly sales dashboard with charts" --response task-keyword-match.json` ranks likely tasks for one English keyword query.
  - `--doctor-request` lints a request, resolves source-backed authored inputs, preflights existing workbook-source access, and returns a machine-readable diagnostics report with every independently provable blocking problem it can isolate safely without mutating a workbook, including multiple malformed steps in one pass.
  - `--doctor-request --request request.json --response doctor-report.json` saves that diagnostics report to disk when stdout is not the right transport.
- Want Java instead of raw JSON: [JAVA_AUTHORING.md](./JAVA_AUTHORING.md) and
  [../examples/java-authoring-workflow.java](../examples/java-authoring-workflow.java)
- Want a no-save health check: [workbook-health-request.json](../examples/workbook-health-request.json)
- Want a copy-sheet maintenance example: [sheet-maintenance-request.json](../examples/sheet-maintenance-request.json)
- Want short copy-paste patterns: [QUICK_REFERENCE.md](./QUICK_REFERENCE.md)
- Want the full field list: [OPERATIONS.md](./OPERATIONS.md)
- Want failure handling: [ERRORS.md](./ERRORS.md)

## Common First-Run Mistakes

- Using `.xls`, `.xlsm`, or `.xlsb` instead of `.xlsx`
- Mixing up path roots: `--response`, `--execution-root`, and `--temp-root` follow the shell
  working directory, while relative paths inside a `--request` file follow that request file's
  directory, and stdin-driven requests use the explicit `--execution-root`
- Ignoring stderr after a failed `--response` run: GridGrind prints one line naming the written response or doctor-report file so the structured failure payload is easy to find
- Expecting GridGrind to save a workbook after a failed run

For hard limits and supported boundaries, see [LIMITATIONS.md](./LIMITATIONS.md).
