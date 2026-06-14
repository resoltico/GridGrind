# GridGrind — .xlsx workbook automation from a JSON request

GridGrind is a `.xlsx` automation engine. Describe workbook work as a JSON request — create sheets,
write cells, build tables, assert results, read facts back. GridGrind runs the whole plan and
returns a structured JSON response. If workbook execution fails, no workbook is saved.

The usual alternative is a mix of libraries, helper scripts, and post-write checks that run after
the file is already saved — with no clean rollback when something fails mid-run. GridGrind replaces
that split with one atomic pass: request in, result out, workbook written only when every step
succeeds.

- Write `.xlsx` workbooks from JSON: sheets, cells, styles, tables, formulas, charts, drawings
- Read facts back in the same plan: cell values, sheet layout, health analysis, pivot data
- Assert workbook state mid-run — a failed assertion stops the plan before saving
- Run from Docker, the packaged `gridgrind` launcher, or a self-contained JAR, against new
  workbooks or existing `.xlsx` files

## Where it fits

Good fit:
- Recurring `.xlsx` workbook jobs that should run the same way each time — filing, updating,
  checking, extracting
- Automation and agent pipelines that create or maintain Excel files without a UI
- Workbook health checks and fact extraction without saving a file
- Environments without Excel — Linux containers, CI pipelines, server-side workflows

Skip it when:
- You need `.xls`, `.xlsm`, or `.xlsb` — GridGrind handles `.xlsx` only
- Your work is truly one-off and hand-writing JSON adds more friction than it saves
- You need interactive formula recalculation during editing (GridGrind evaluates on request)

## Quick Start

From a repository checkout, the shortest reliable path is:

```bash
./gradlew :cli:installShadowDist
./cli/build/install/gridgrind/bin/gridgrind --help
./cli/build/install/gridgrind/bin/gridgrind --print-example --lookup BUDGET --response budget-request.json
./cli/build/install/gridgrind/bin/gridgrind --doctor-request --request budget-request.json --response doctor-report.json
./cli/build/install/gridgrind/bin/gridgrind --request budget-request.json --response response.json
```

In the snippets below, read `gridgrind` as the entry point you actually have: the packaged repo
launcher (`./cli/build/install/gridgrind/bin/gridgrind`), the `bin/gridgrind` launcher from a
release archive, or `java -jar gridgrind.jar` if you downloaded the standalone JAR.

For first contact, prefer `--request <path>` over stdin. Stdin-driven execution and doctoring
require `--execution-root <path>` so request-owned paths resolve from one explicit invocation root.

If you want the repository JAR surface directly, run `./gradlew :cli:shadowJar` and replace
`gridgrind` below with `java -jar cli/build/libs/gridgrind.jar`.

If you already have a release artifact, use the unpacked `bin/gridgrind` launcher from the zip/tar
distribution, or invoke the standalone `gridgrind.jar` with `java -jar gridgrind.jar`.

If you want the container surface from a repository checkout, the root Dockerfile now builds the
packaged runtime image on its own:

```bash
docker buildx build --load -t gridgrind-local .
docker run --rm gridgrind-local --help
```

Fast Docker first-contact:

```bash
docker run --pull=always --rm ghcr.io/resoltico/gridgrind:latest --help
```

## Find The Right Starting Point

GridGrind can print valid starting material instead of making you invent request shape by hand:

```bash
gridgrind --print-request-template --response request.json
gridgrind --print-example-catalog --response examples.json
gridgrind --print-task-catalog --response tasks.json
gridgrind --print-task-plan --lookup DASHBOARD --response dashboard-request.json
gridgrind --print-task-keyword-match --query "monthly sales dashboard" --response task-match.json
gridgrind --print-protocol-catalog --response protocol-index.json
gridgrind --print-protocol-catalog --search pivot --response pivot-search.json
gridgrind --print-protocol-catalog --lookup mutationActionTypes:SET_CELL --response set-cell.json
gridgrind --print-protocol-catalog --full --response protocol-catalog.json
```

The example and task catalogs publish `requestFileName`, `workspaceMode`, and
`requiredWorkspacePaths`, so you can tell whether a printed request is self-contained before you
try to run it.

The bare `--print-protocol-catalog` output is the compact first-contact index. Use
`--print-protocol-catalog --search <text>` when you know the concept but not the exact id, follow
up with `--print-protocol-catalog --lookup <group>:<id>` when you want one full authoritative
entry, and use `--print-protocol-catalog --full` only when you need the entire machine-readable
catalog in one payload.

Use `--help` for the short synopsis, `--help-protocol` for the authoritative CLI and request
contract, and `--help-guidance` for workflow-oriented help.

## One Request, One Result

A single JSON request describes every step: create a sheet, write cells, assert workbook state,
read facts back, and save. GridGrind executes the steps in order and writes the file only when
every step succeeds. If an assertion fails or any step errors, no workbook is saved.

The top-level envelope is always explicit: `protocolVersion`, `source`, `persistence`,
`execution`, `formulaEnvironment`, and ordered `steps`. Steps can mix mutation, assertion, and
inspection in the same plan.

The safest way to start is to ask GridGrind to emit a valid request for you:

```bash
gridgrind --print-request-template --response request.json
gridgrind --print-example --lookup BUDGET --response budget-request.json
gridgrind --print-task-plan --lookup DASHBOARD --response dashboard-request.json
```

## Documentation

- [Full docs index](docs/INDEX.md) — every reference file organized by topic
- [First run guide](docs/QUICK_START.md) — first successful run, Docker or JAR
- [Snippets](docs/QUICK_REFERENCE.md) — copy-paste request patterns
- [Java authoring](docs/JAVA_AUTHORING.md) — build requests from Java instead of JSON
- [Operations reference](docs/OPERATIONS.md) — every field and operation
- [Examples](examples/) — shipped request files plus companion assets for the asset-backed cases

## Legal

GridGrind is MIT-licensed. Its executable JAR bundles third-party components under Apache 2.0,
BSD 2-Clause, BSD 3-Clause, and EDL 1.0 licenses. See [NOTICE](NOTICE) for the complete
attribution list and [PATENTS.md](PATENTS.md) for patent considerations.

[LICENSE](LICENSE) | [NOTICE](NOTICE) | [PATENTS.md](PATENTS.md) | [LICENSE-APACHE-2.0](LICENSE-APACHE-2.0) | [LICENSE-BSD-2-CLAUSE](LICENSE-BSD-2-CLAUSE) | [LICENSE-BSD-3-CLAUSE](LICENSE-BSD-3-CLAUSE) | [LICENSE-EDL-1.0](LICENSE-EDL-1.0)
