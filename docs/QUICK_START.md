---
afad: "3.5"
version: "0.49.0"
domain: QUICK_START
updated: "2026-04-21"
route:
  keywords: [gridgrind, quick start, first run, docker, jar, xlsx, example, response]
  questions: ["how do i do a first run with gridgrind", "what is the fastest way to try gridgrind", "how do i run the shipped examples", "how do i get my first successful gridgrind run"]
---

# Quick Start

**Purpose**: Get to a first successful GridGrind run with the least setup and the least guesswork.
**Best starting point**: the shipped [budget example](../examples/budget-request.json).

---

## What You Need

- A GridGrind runtime: Docker image or the release JAR
- A working directory where GridGrind can read the request file and write the response file
- One example request to start from: [budget-request.json](../examples/budget-request.json) is the simplest first choice

GridGrind supports `.xlsx` workbooks only.

If you are starting from the release JAR alone, copy `examples/budget-request.json` into your
working directory first so the request path in the commands below exists.

## Pick One Run Path

### Docker

If Docker is the easiest path on your machine, pull the latest image:

```bash
docker pull ghcr.io/resoltico/gridgrind:latest
```

### Release JAR

If you want the standalone JAR, download it from the
[latest release](https://github.com/resoltico/GridGrind/releases/latest).

The JAR path requires Java 26.

## First Successful Run

Use the shipped [budget example](../examples/budget-request.json) for the first pass. It writes a
sample workbook and a JSON response, so you can see both the output file and the run result.

### Docker Example

Run from the repository root or another directory that contains the example request:

```bash
docker run -i \
  -v "$(pwd)":/workdir \
  -w /workdir \
  ghcr.io/resoltico/gridgrind:latest \
  --request examples/budget-request.json \
  --response response.json
```

### JAR Example

Replace `gridgrind.jar` with the downloaded JAR filename if it differs on your machine.

```bash
java -jar gridgrind.jar \
  --request examples/budget-request.json \
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
