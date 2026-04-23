---
afad: "3.5"
version: "0.55.0"
domain: QUICK_START
updated: "2026-04-22"
route:
  keywords: [gridgrind, quick start, first run, docker, jar, xlsx, example, response]
  questions: ["how do i do a first run with gridgrind", "what is the fastest way to try gridgrind", "how do i run the shipped examples", "how do i get my first successful gridgrind run"]
---

# Quick Start

**Purpose**: Get to a first successful GridGrind run with the least setup and the least guesswork.
**Best starting point**: generate the built-in `BUDGET` example directly from the artifact with `--print-example BUDGET`. If you are already in a repo checkout, the matching JSON also lives at [../examples/budget-request.json](../examples/budget-request.json).

---

## What You Need

- A GridGrind runtime: Docker image or the release JAR
- A working directory where GridGrind can read the request file and write the response file
- One example request to start from: the built-in `BUDGET` example emitted by `--print-example BUDGET`, or [budget-request.json](../examples/budget-request.json) when you are already in a repo checkout

GridGrind supports `.xlsx` workbooks only.

If you are starting from the release artifact alone, generate the request into your working
directory first so the later `--request` path already exists: `gridgrind --print-example BUDGET > budget-request.json`.

## Pick One Run Path

### Docker

If Docker is the easiest path on your machine, pull the latest image:

```bash
docker pull ghcr.io/resoltico/gridgrind:latest
```

The published image already includes the font stack required for signature-line preview
generation, so signature-line requests work in Docker without extra image customization.

### Release JAR

If you want the standalone JAR, download it from the
[latest release](https://github.com/resoltico/GridGrind/releases/latest).

The JAR path requires Java 26.

## First Successful Run

Use the built-in `BUDGET` example for the first pass. It writes a sample workbook and a JSON
response, so you can see both the output file and the run result. If you are already in a repo
checkout, [budget-request.json](../examples/budget-request.json) is the matching checked-in copy.

### Docker Example

Generate the built-in request once, then run it from the current directory:

```bash
docker run --rm ghcr.io/resoltico/gridgrind:latest --print-example BUDGET \
  > budget-request.json

docker run -i \
  -v "$(pwd)":/workdir \
  -w /workdir \
  ghcr.io/resoltico/gridgrind:latest \
  --request budget-request.json \
  --response response.json
```

### JAR Example

Replace `gridgrind.jar` with the downloaded JAR filename if it differs on your machine.

```bash
java -jar gridgrind.jar --print-example BUDGET > budget-request.json

java -jar gridgrind.jar \
  --request budget-request.json \
  --response response.json
```

## What To Look For

After a successful run:

- `response.json` should report `status: "SUCCESS"`
- the budget example writes its workbook to
  `cli/build/generated-workbooks/gridgrind-budget.xlsx`
- if the run fails, GridGrind returns a structured error response instead of saving a partial workbook

## Good Second Steps

- Want GridGrind to explain itself from the artifact instead of from prose:
  - `--print-task-catalog` lists the contract-owned high-level office-work tasks.
  - `--print-task-plan DASHBOARD` emits a starter request scaffold for one task.
  - `--print-goal-plan "monthly sales dashboard with charts"` ranks likely tasks for one freeform goal.
  - `--doctor-request` lints a request and returns a machine-readable diagnostics report without opening or mutating a workbook.
- Want a no-save health check: [workbook-health-request.json](../examples/workbook-health-request.json)
- Want short copy-paste patterns: [QUICK_REFERENCE.md](./QUICK_REFERENCE.md)
- Want the full field list: [OPERATIONS.md](./OPERATIONS.md)
- Want failure handling: [ERRORS.md](./ERRORS.md)

## Common First-Run Mistakes

- Using `.xls`, `.xlsm`, or `.xlsb` instead of `.xlsx`
- Running from the wrong working directory, so the request path does not resolve
- Expecting GridGrind to save a workbook after a failed run

For hard limits and supported boundaries, see [LIMITATIONS.md](./LIMITATIONS.md).
