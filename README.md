# GridGrind — .xlsx workbook automation from a JSON request

GridGrind is a `.xlsx` automation engine. Describe workbook work as a JSON request — create sheets,
write cells, build tables, assert results, read facts back. GridGrind runs the whole plan and
returns a structured JSON response. Failures prevent persistence: assertions fail fast by default,
or `COLLECT` completes the terminal verification phase and preserves evidence for every failed assertion.

The usual alternative is a mix of libraries, helper scripts, and post-write checks that run after
the file is already saved — with no clean rollback when something fails mid-run. GridGrind replaces
that split with one atomic pass: request in, result out, workbook written only when every step
succeeds.

- Write `.xlsx` workbooks from JSON: sheets, cells, styles, tables, formulas, charts, drawings
- Read facts back in the same plan: cell values, sheet layout, health analysis, pivot data
- Assert workbook state mid-run — fail fast by default, or collect a terminal assertion phase before saving
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
export PATH="$(pwd)/cli/build/install/gridgrind/bin:$PATH"
gridgrind --help
gridgrind --print-recipe --lookup BUDGET --response budget-request.json
mkdir -p generated-workbooks
gridgrind --doctor-request --request budget-request.json --response doctor-report.json
gridgrind --request budget-request.json --response response.json
```

The rest of this README uses `gridgrind` for the active entry point. From a repository checkout,
the `export PATH=...` line above points that name at the packaged launcher. From a Gradle-generated
ZIP or TAR launcher distribution, add its `bin/` directory to `PATH`; from the standalone GitHub
release JAR, replace `gridgrind` with `java -jar gridgrind.jar`.

Without `--response`, GridGrind writes one primary JSON payload to stdout. A command rejected
before workbook execution uses `CommandError` with `status: "REJECTED"`; execution uses
`WorkbookResult` with `status: "SUCCEEDED"` or `"FAILED"`. With `--response`, that payload goes
to the requested file instead. For execution, GridGrind reserves a new no-follow response file
after request validation and before input binding or workbook work; an existing, directory, or
unwritable destination prevents execution and emits a typed transport notice on stderr. If a
reserved response file later fails while bytes are written, the already-rendered primary payload
goes to stdout unchanged when stdout is available and stderr contains one small transport-only JSON
notice. GridGrind never moves a primary payload to stderr or overwrites an existing response file.

For first contact, prefer `--request <path>` over stdin. Stdin-driven execution and doctoring
require `--execution-root <path>` so request-owned paths resolve from one explicit invocation root.
`--doctor-request` returns every independently provable blocking problem it can isolate safely before any workbook mutation or save attempt. Request intake reports invalid UTF-8, duplicate keys, unknown fields, omitted required fields, explicit nulls, malformed scalar values, missing or unknown type discriminators, and constructor-level field validation failures together while retaining valid sibling fragments. Bound steps then receive their operation-and-target contract checks, and a fully bound plan also batches independent source and authored-input preflight findings in the same doctor report. Normal `--request` performs the same intake and static validation before execution; static failures are rejected before workbook access, while any later preflight failure completes zero steps and persists no workbook.

If you want the repository JAR surface directly, run `./gradlew :cli:shadowJar` and invoke
`java -jar cli/build/libs/gridgrind.jar ...` instead of `gridgrind ...`.

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

For file-producing container runs, mount your workspace at `/work`, pass
`--user "$(id -u):$(id -g)"` on ordinary bind mounts so written files stay owned by your host
user, and let the image's prepared `WORKDIR` resolve relative request and response paths from
there; omit `--user` only when Docker Desktop or a rootless runtime already remaps bind-mount
ownership for you.

```bash
docker run --pull=always --rm -i \
  --user "$(id -u):$(id -g)" \
  -v "$(pwd)":/work \
  ghcr.io/resoltico/gridgrind:latest \
  --request request.json \
  --response response.json
```

## Find The Right Starting Point

GridGrind can print valid starting material instead of making you invent request shape by hand:

```bash
gridgrind --print-request-template --response request.json
gridgrind --print-recipe-catalog --response recipes.json
gridgrind --print-recipe-catalog --lookup DASHBOARD --response dashboard-detail.json
gridgrind --print-recipe --lookup DASHBOARD --response dashboard-request.json
gridgrind --materialize-recipe --lookup CUSTOM_XML --workspace custom-xml-workspace
gridgrind --print-recipe-keyword-match --query "monthly sales dashboard" --response task-match.json
gridgrind --print-protocol-catalog --response protocol-index.json
gridgrind --print-protocol-catalog --search pivot --response pivot-search.json
gridgrind --print-protocol-catalog --lookup mutationActionTypes:SET_CELL --response set-cell.json
gridgrind --print-protocol-catalog --lookup plainTypes:cellReadProjectionType --response cell-read-projection.json
```

The compact recipe catalog publishes `requestFileName`, `advisory`, and `requiredWorkspacePaths`, while `--print-recipe-catalog --lookup <id>` adds the exact runnable request profile for that recipe. Print self-contained recipes with `--print-recipe`; materialize asset-backed recipes with `--materialize-recipe --workspace <new-directory>`, which atomically writes the request and declared assets together. `VERBOSE` execution streams compact JSONL progress to stderr while its primary result remains on stdout or the requested response file. Shipped save-producing examples already use `SAVE_AS.ifExists=REPLACE`, so rerunning them does not depend on cleaning the workspace first.

The bare `--print-protocol-catalog` output is the compact first-contact index. Use
`--print-protocol-catalog --search <text>` when you know the concept but not the exact id, follow
up with `--print-protocol-catalog --lookup <group>:<id>` when you want one authoritative payload,
and keep discovery scoped instead of expecting one monolithic full-catalog dump. Field descriptors
publish enforced scalar constraints and operations publish applicable preconditions, so tools can
validate request material without extracting rules from prose. Machine-readable request-template,
discovery, doctor, and execution payloads are compact JSON by default; add `--pretty` when you want
indented JSON instead.

Use `--help` for the short synopsis, `--help-protocol` for the authoritative CLI and request
contract, and `--help-guidance` for workflow-oriented help.

## One Request, One Result

A single JSON request describes every step: create a sheet, write cells, assert workbook state,
read facts back, and save. GridGrind executes the steps in order and writes the file only when
every step succeeds. Assertions fail fast by default; `execution.assertionMode=COLLECT` instead
runs every assertion in a terminal verification phase, retains complete evidence for each failed
assertion in `assertions[]`, and returns the first canonical failure. Any assertion failure or step
error prevents persistence.

The smallest valid top-level envelope is `protocolVersion`, `source`, `persistence`, and ordered
`steps`. `execution` and `formulaEnvironment` are optional when you want the default
`FULL_XSSF` / `SUMMARY` / `DO_NOT_CALCULATE` execution path and the empty evaluator environment.
Steps can mix mutation, assertion, and inspection in the same plan. Every response also carries one
top-level `persistence` outcome so callers can see both the requested save mode and whether a file
was actually written. `FORMULA` and `RAW_FORMULA` values are OOXML formula bodies and must not
begin with `=`; normal `FORMULA` values are validated when their mutation executes, while
`RAW_FORMULA` remains the explicit opaque path for newer Excel syntax. `SAVE_AS` requires an
explicit `ifExists=REJECT|REPLACE` choice; use `REPLACE` when you want rerunnable create-or-replace
output.

The safest way to start is to ask GridGrind to emit a valid request for you:

```bash
gridgrind --print-request-template --response request.json
gridgrind --print-recipe --lookup BUDGET --response budget-request.json
gridgrind --print-recipe --lookup DASHBOARD --response dashboard-request.json
```

## Documentation

- [Full docs index](docs/INDEX.md) — every reference file organized by topic
- [First run guide](docs/QUICK_START.md) — first successful run from the packaged launcher, Docker, or JAR
- [Snippets](docs/QUICK_REFERENCE.md) — copy-paste request patterns
- [Java authoring](docs/JAVA_AUTHORING.md) — build requests from Java instead of JSON
- [Operations reference](docs/OPERATIONS.md) — every field and operation
- [Examples](examples/) — shipped request files plus companion assets for the asset-backed cases

## Legal

GridGrind's own source code is MIT-licensed. The executable JAR also bundles separately licensed
components under Apache 2.0, MIT-style, Boost 1.0, BSD 2-Clause, and BSD 3-Clause terms; the Jakarta
API terms are formally the Eclipse Distribution License v1.0, whose SPDX identifier is
`BSD-3-Clause`. The standalone GitHub release JAR embeds the legal files, while Gradle-generated ZIP
and TAR launcher distributions also place them beside the launcher. `gridgrind --license` prints
the GridGrind license, third-party notices, and applicable dependency license texts.

Public source archives additionally include the Apache-2.0 Gradle wrapper. A byte-identical Apache
POI Custom XML test workbook is present both in the source tree and in the executable JAR's packaged
recipe assets; [NOTICE](NOTICE) identifies its provenance, hash, and exact distribution scope.

The published container additionally contains an Azul Zulu OpenJDK runtime and operating-system
and font packages under their own terms. GridGrind's legal files are available at
`/usr/share/doc/gridgrind`, while the runtime and package legal materials remain in their standard
JRE and `/usr/share/doc` locations; published images also include provenance and an SBOM
attestation. See [NOTICE](NOTICE) for the audited component inventory and attributions, and
[PATENTS.md](PATENTS.md) for the deliberately limited patent statement.

[LICENSE](LICENSE) | [NOTICE](NOTICE) | [PATENTS.md](PATENTS.md) | [LICENSE-APACHE-2.0](LICENSE-APACHE-2.0) | [LICENSE-BSD-2-CLAUSE](LICENSE-BSD-2-CLAUSE) | [LICENSE-BSD-3-CLAUSE](LICENSE-BSD-3-CLAUSE)
