# Changelog

Notable changes to this project are documented in this file. The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.60.0] - 2026-04-28

### Changed

- GridGrind now ships a committed contributor devcontainer based on a pinned glibc image plus
  Azul Zulu 26, and the developer docs now treat that containerized workflow as the preferred
  contributor path while keeping host-native Java guidance accurate as a fallback.
- The shared Gradle build logic now uses Kotlin `2.4.0-Beta2` on Gradle `9.5.0-rc-4` and emits
  JVM 26 bytecode directly, so the build no longer carries a separate JVM 25 exception below the
  repository's Java 26 baseline.
- The canonical `./check.sh` gate now runs with an isolated repo-scoped `GRADLE_USER_HOME`,
  always forces `--no-daemon`, and shares one repo-wide verification lock with devcontainer,
  Docker-smoke, and Jazzer wrapper entrypoints so local full verification no longer overlaps
  through shared daemon or cache state by accident.
- `./check.sh` now derives its fixed five-stage inventory, Stage 4 release-surface shell
  regression list, and usage text from `scripts/check-stage-contract.sh`, and a dedicated shell
  regression now guards that canonical owner so the root gate cannot drift back into parallel
  stage definitions.
- Discovery and planning JSON now omit absent optional fields instead of emitting explicit
  `null` placeholders. `--print-protocol-catalog`, `--print-task-catalog`,
  `--print-task-plan`, `--print-goal-plan`, and `--doctor-request` stay schema-compatible while
  becoming easier for agents and shell tooling to consume directly.
- Request and response JSON now omit absent optional fields too, including `--print-example`
  output, checked-in `examples/*.json`, and normal execution responses, so the full public JSON
  surface no longer uses explicit `null` placeholders as a control protocol.
- Workbook-analysis finding codes and severities now come from one shared
  `excel-foundation` owner instead of parallel copies in engine and contract modules, so the
  published vocabulary no longer depends on lockstep duplicate enums.
- The limitations registry now distinguishes product-enforced limits from upstream reference
  ceilings instead of implying that every documented Excel or POI ceiling is surfaced through the
  same runtime/help/catalog propagation path.
- Canonical request plans now keep default execution and formula-environment state as normalized
  objects in memory instead of collapsing those sections back to `null`, so authoring, executor,
  and JSON emission all share one default model.
- Request-doctor and calculation telemetry contracts now keep optional summary, problem,
  preflight, and message state as explicit `Optional`-backed fields in memory instead of raw
  `null` padding, so Java authoring, CLI fallback handling, executor telemetry, and tests all
  share the same absence model as the public JSON wire shape.
- Jazzer's deterministic coverage gate now treats the extracted `OperationSequence*` generator
  helpers as one exclusion family, so splitting the fuzz-workflow generator into smaller seams no
  longer breaks Stage 2 verification just because the coverage scope list drifted.

### Fixed

- CI now validates the committed devcontainer surface in addition to the whole-repo deterministic
  gate, so the preferred contributor environment cannot silently drift away from the documented
  Java, editor-routing, and container-base contract.
- Release CI no longer strands branch protection on stale check names while the release verifiers
  silently ignore the committed devcontainer job. The `CI` workflow now keeps `Check` and
  `Docker smoke` as the required contexts, exposes `Contributor devcontainer` as a separate job,
  and the tag/merge release verifiers now treat that contributor-environment job as
  release-blocking too.
- `./check.sh` progress monitoring no longer assumes BSD `stat`. The root gate now sources a
  portable file-size helper with BSD, GNU, and generic fallback behavior, and a dedicated shell
  regression keeps the Linux/macOS seam from breaking release CI again.
- The CLI-contract regression no longer pushes full fixture payloads through huge process
  environments. It now writes case fixtures under `tmp/` and verifies the contract through file
  paths, avoiding Linux `Argument list too long` failures in release CI.
- The engine now names every private Apache POI contract it still depends on for relation
  removal, sheet-clone hyperlink preparation, workbook picture-catalog synchronization, and
  gradient-fill registry access. Compatibility tests and runtime failures now point at the exact
  broken POI seam and affected feature instead of failing later with generic reflection errors.
- `--print-protocol-catalog` no longer routes full-catalog, ranked-search, and exact-entry lookup
  through one nullable union command model. The CLI now uses dedicated typed variants internally,
  so discovery behavior is explicit instead of being inferred from null combinations.
- Protocol-catalog target-selector metadata no longer relies on internal null-sentinel helper
  returns. Derived selector rules now flow through typed `Optional` seams, and exact entry lookups
  stop printing `targetSelectorRule: null` noise for operations that do not need a rule.
- Release-surface shell regressions no longer validate the CLI contract against hand-maintained
  fake help/catalog/task/goal/doctor blobs. The verifier tests now derive their baseline payloads
  from the built GridGrind CLI artifact and mutate only the field under test.
- Current agent and Jazzer architecture guidance no longer teaches the deleted top-level
  `protocol` module, and a build-failing architecture audit now guards the live
  `cli -> executor -> contract` plus `executor -> engine` graph.
- The success/failure response seam no longer depends on nested null-padding helpers for cell and
  problem-context variants. `CellReport` and `ProblemContext` now live as dedicated top-level
  contract types, and a build-failing seam audit guards `GridGrindResponse.java` against growing
  back into an unbounded god-file.
- Build-failing seam audits now also guard `ProblemContext.java`,
  `OperationSequenceModel.java`, and `OperationSequenceValueFactory.java` after splitting their
  selector/assertion and chart-generation sub-foundations into dedicated helpers, so contract and
  Jazzer boundary theory no longer depends on a handful of oversized god-files.
- Chart clone-preparation helpers no longer use raw `null` to represent missing formula-node text
  payloads. The rewrite seam now uses `Optional` internally and preserves the same public failure
  when POI loses the node text during restore.

## [0.59.0] - 2026-04-25

### Changed

- Structured success and failure response construction now uses explicit protocol factories for
  synthetic journal-backed results, instead of public convenience constructors that inferred
  omitted execution details from `null`.
- Presence assertions are now family-specific on the public wire surface. `EXPECT_PRESENT` and
  `EXPECT_ABSENT` have been replaced by `EXPECT_NAMED_RANGE_PRESENT`,
  `EXPECT_NAMED_RANGE_ABSENT`, `EXPECT_TABLE_PRESENT`, `EXPECT_TABLE_ABSENT`,
  `EXPECT_PIVOT_TABLE_PRESENT`, `EXPECT_PIVOT_TABLE_ABSENT`, `EXPECT_CHART_PRESENT`, and
  `EXPECT_CHART_ABSENT`, so selector intent is explicit in the authored request instead of being
  recovered from ambiguous shared selector ids.
- Generated CLI help and the machine-readable CLI catalog now publish first-contact workflows for
  discovery, drafting, doctoring, and execution, instead of presenting the public surface only as
  a flat flag inventory.
- Generated CLI help now teaches the save-then-reopen workflow explicitly too, including
  `source.type=EXISTING` for reopening an `.xlsx` workbook from disk and the matching
  `source.path` file-workflow wording, so first-contact operators do not have to infer the
  existing-workbook discriminator from deeper protocol material.
- Generated CLI help, the quick reference, and the request/execution reference now call out that
  every non-empty step needs a caller-defined `stepId`, so first-contact request authors do not
  have to discover that requirement only from a shape error after writing their first mutation or
  inspection.

### Fixed

- Engine and executor helper seams no longer rely on test-only deep reflection or null-sentinel
  returns for key optional facts. Residual coverage now exercises package-visible seams directly,
  autofilter sort-state readback uses typed absence internally, and executor warning, diagnostic,
  and analysis-severity helpers no longer encode "missing" with raw `null` in business logic.
- Literal cell writes now actually replace existing formula cells instead of only updating POI's
  cached formula result. `SET_CELL` and `SET_RANGE` clear the old formula first, so writing a
  number, text, boolean, blank, date, or date-time value no longer leaves the prior formula alive
  under the new cached display value.
- `INSERT_ROWS` and `INSERT_COLUMNS` no longer fail just because the target sheet already owns
  data validations. GridGrind now snapshots those validation structures before the structural
  insert, shifts their covered ranges and supported formulas authoritatively, and reapplies them
  after the insert so validated workbooks can be extended without a manual clear-and-rebuild pass.
- `COPY_SHEET` no longer crashes when Apache POI tries to clone charts whose data sources are
  authored through workbook or sheet defined names. GridGrind now normalizes those chart formulas
  to explicit area references just for the clone pass, restores the source-sheet chart formulas
  immediately afterward, and leaves the copied chart bound to the copied sheet instead of failing
  with a negative-column reference error.
- Failure diagnostics now require an explicit `problem.causes[*].stage` value, so structured cause
  entries always preserve the pipeline stage that classified the failure.
- Contract-exposed foundation enum tokens and gradient-fill type tokens are now pinned by
  regression coverage, so accidental wire-vocabulary renames fail verification instead of
  silently drifting the published protocol surface.
- `./check.sh` Stage 4 documentation and usage text now list `scripts/test-verify-cli-contract.sh`
  alongside the other release-surface shell regressions that the script already executes.
- `--doctor-request` now accepts `--response <path>` just like normal execution, so machine-
  readable doctor reports can be written to files in both fat-JAR and Docker runs instead of
  forcing stdout-only handling.
- `--response <path>` now works consistently across discovery and informational CLI commands too,
  including `--help`, `--version`, `--print-request-template`, `--print-example`,
  `--print-task-catalog`, `--print-task-plan`, `--print-goal-plan`, and
  `--print-protocol-catalog`, with the same parent-directory creation and structured fallback
  behavior already used by execution and doctoring.
- Missing polymorphic `type` discriminators in request JSON no longer crash the parser's public
  error-reporting path. Requests that omit a step assertion, action, query, source, or other
  subtype `type` field now fail deterministically with the product-owned
  `Missing required field 'type'` contract instead of surfacing a null-discriminator
  `NullPointerException`.
- Public Docker-first guidance now calls out that a plain `docker run ...:latest` can reuse a
  stale local `latest` tag. Documentation now tells operators to `docker pull` first or use
  `--pull=always` when they need the registry's current `latest`.
- User-facing Docker-first copy-paste examples now use the freshness-safe form too. Documentation
  no longer warns about stale `latest` tags in prose while still demonstrating plain
  `docker run ...:latest` as the first command a new operator is likely to copy.
- First-contact request-shape failures now teach the intended replacement surface instead of only
  echoing an unknown discriminator token. `source.type=FILE` now points operators at
  `source.type=EXISTING`, the removed ambiguous assertion ids point at the new family-specific
  assertion ids, and `--print-example` now suggests the stable upper-case example id when the
  authored token matches a shipped example file stem such as `chart-request`.

## [0.58.0] - 2026-04-24

### Changed

- Public docs now distinguish built-in examples that are fully self-contained in a blank artifact
  workspace from the three repo-asset-backed examples (`CUSTOM_XML`, `SOURCE_BACKED_INPUT`, and
  `PACKAGE_SECURITY_INSPECTION`) that require copied `examples/` assets. Runtime example tests now
  cover that portability split directly instead of leaving it implied by prose.
- Public docs and generated help now describe the doctor surface more precisely:
  `--doctor-request` validates request shape, source-backed authored input resolution, and
  existing-workbook source accessibility up front, while still leaving workbook mutation to the
  real execution path.
- Built-in example discovery now publishes machine-readable workspace requirements through
  `workspaceMode` and `requiredPaths`, so operators and agents can tell which examples need copied
  repository assets without relying on prose or trial execution.
- The release protocol now explicitly treats mounted or removable primary checkouts that break
  Gradle locking or stall preflight as problematic release hosts and directs operators onto a clean
  local worktree before any release build, instead of leaving that failure mode implicit.

### Fixed

- `SAVE_AS` persistence now creates missing parent directories consistently on both the normal
  XSSF path and the `STREAMING_WRITE` path, matching the documented file-workflow contract.
- `--doctor-request` no longer reports some requests as valid when they would immediately fail
  during input resolution or workbook opening at execution time. Source-backed authored input
  loading and existing-workbook accessibility are now preflighted in doctor mode as well.
- Bare `help` now works as a first-class alias for `--help` in the packaged CLI instead of failing
  as an unknown argument.
- Non-step CLI failures now preserve their real problem classification in the synthetic execution
  journal instead of always backfilling `INTERNAL_ERROR`.
- Root-project formatting checks no longer race concurrent compilation by scanning a mutating
  checkout during `./check.sh`; the repository-wide Spotless project-file pass now runs against a
  stable tree before build outputs are produced.

## [0.57.0] - 2026-04-24

### Fixed

- Running the packaged CLI with no arguments from an interactive terminal no longer hangs silently
  while waiting for stdin. `gridgrind` and `java -jar gridgrind.jar` now print the same help text
  as `--help` and exit with code `0`, including cases where stdin is interactive but stdout is
  redirected. The release-surface CLI verifier now asserts that interactive no-arg behavior under
  a real PTY so this path cannot regress quietly again.
- The Java authoring surface now stops cleanly at the canonical `WorkbookPlan` boundary instead of
  pulling executor runtime concerns into the `authoring-java` module. `dev.erst.gridgrind.authoring`
  now depends transitively on `contract`, the fluent API no longer exposes `run(...)`, raw
  request DTOs, or response-report types as intended public authoring vocabulary, and the
  shipped Java example now executes in process through an explicit executor call.
- Goal planning is now less noisy for agents and operators. `--print-goal-plan` no longer keeps
  low-signal candidates alive solely because one capability summary happened to share a word with
  the goal, and `suggestedIntentTags` now come from the strongest matching tasks instead of from
  the entire catalog.
- Protocol-catalog field-group metadata is no longer collapsed into one giant nested/plain support
  file. The nested and plain descriptor registries now live in dedicated support types again, with
  the thin facade left in place only as the local assembly seam.
- Release-surface shell regressions no longer duplicate the same fake help/catalog/task/doctor
  payloads in multiple scripts. The shared fake public-contract fixtures now live in one sourced
  shell helper so verifier updates have one maintenance point instead of two drifting copies.

## [0.56.0] - 2026-04-23

### Changed

- Task discovery now ships broader runnable starter scaffolds instead of thin placeholders.
  `--print-task-catalog`, `--print-task-plan`, and `--print-goal-plan` now cover tabular
  reports, dashboards, data-entry workflows, workbook audits, custom XML workflows, pivot
  reports, drawing/signature workflows, and workbook-maintenance flows with non-empty starter
  plans.

### Fixed

- CLI request-path handling is now portable and internally consistent. When the CLI reads a
  request via `--request`, relative request-owned paths now resolve from the request file
  directory instead of the shell working directory, including `source.path`, persistence paths,
  source-backed `UTF8_FILE` / `FILE` payloads, formula external workbook bindings, and OOXML
  signing keystore paths.
- CLI discovery commands now reject stray trailing flags consistently instead of silently
  accepting malformed invocations. `--help`, `--version`, `--license`, `--print-example`,
  `--print-task-catalog`, `--print-task-plan`, `--print-goal-plan`, `--print-request-template`,
  and `--print-protocol-catalog` now fail fast on extra arguments, and `--doctor-request` no
  longer ignores unknown trailing flags.
- Public docs, help text, and shipped examples now match the live discovery and path-resolution
  surface. The request/execution reference, quick-start flow, quick reference, README, operations
  index, and generated example set now document request-file-rooted relative paths, the broader
  task planner surface, and the new sheet-maintenance example.
- Release-surface shell validation no longer assumes every verification checkout is a live Git
  worktree. The Jazzer public-surface regression still verifies Git tracking when `.git` is
  present, but copied validation checkouts now pass based on the actual shipped shell surface
  instead of failing on missing repository metadata, and `./check.sh` now truthfully lists that
  Jazzer public-surface regression in its own Stage 4 contract summary.
- The largest new discovery seam no longer lives in one 900-line registry file. Task-definition
  builders are now split by responsibility, and the executor conversion surface remains split
  across workbook, cell, drawing, structured-feature, read-result, and workbook-report seams
  instead of regressing into new god-files.

## [0.55.0] - 2026-04-23

### Fixed

- Copied-sheet picture readback no longer depends on Apache POI leaving cloned drawing relation ids
  intact. GridGrind now repairs copied picture `r:embed` references against the target drawing
  relations after `COPY_SHEET`, picture reads or payload extraction now fail with a clear
  picture-specific integrity error instead of a raw `NullPointerException` when an external
  workbook is malformed, and the reproduced Jazzer `COPY_SHEET` picture crash is now covered by a
  promoted round-trip success seed plus persisted drawing-picture OOXML invariants.
- Malformed picture XML that still has `<pic:blipFill>` but no nested `<a:blip>` no longer leaks
  Apache POI's raw `NullPointerException` through picture inspection or delete flows. GridGrind
  now short-circuits those states as missing image relationships, keeps the delete path safe, and
  covers the blip-less picture seam with focused engine regressions.
- Copied embedded objects now keep their preview-image drawing relations intact through
  `COPY_SHEET` round-trips. GridGrind now repairs the preview media relationship referenced from
  the copied object shape itself, so saved copied worksheets no longer reopen with unresolved
  drawing `a:blip` ids in the embedded-object preview path.
- Copied embedded objects now also keep their worksheet-bound relation ids authoritative through
  `COPY_SHEET`, even when cloned `oleObject` or `objectPr` ids collide with the copied sheet's
  drawing, comment-VML, or other worksheet relations. GridGrind now rewrites those copied ids onto
  safe worksheet-owned relation ids, restores the sheet-to-drawing relation explicitly, treats
  sheet preview refs as image-only during read/delete flows, extends persisted OOXML invariants to
  worksheet `objectPr` preview refs, and promotes the reproduced round-trip artifact as a
  committed success seed.

## [0.54.0] - 2026-04-23

### Fixed

- The public contract no longer depends on the POI-backed engine module for shared Excel domain
  types. GridGrind now ships a dedicated `excel-foundation` module and package for the shared
  `.xlsx` enums, limits, and span/value objects consumed by the contract, engine, executor, and
  tests, so the `contract` module is a real boundary instead of a named-module cycle hiding behind
  `dev.erst.gridgrind.excel`.
- Developer architecture docs and release-surface shell regressions now describe and enforce the
  real six-module product graph. `excel-foundation` is now part of the documented JPMS graph, the
  accepted contract-replacement ADR no longer claims GridGrind is still a five-module system, and
  `./check.sh` no longer fails Stage 4 just because the shared foundation module exists.
- Formula authoring and formula-rewrite paths now go through one engine seam instead of each
  mutation mode setting formulas ad hoc. Streaming append writes, direct cell writes, copied-sheet
  sheet-name rewrites, and copied-table structured-reference rewrites now share consistent error
  handling for authored, rewritten, and scratch validation formulas.
- Protocol discovery no longer forces operators and agents to dump the full catalog and grep it
  externally. `gridgrind --print-protocol-catalog --search <text>` now returns ranked
  machine-readable matches across ids, qualified ids, groups, and summaries, and
  `--print-protocol-catalog` now rejects stray trailing flags instead of silently ignoring them.
- Copy-sheet table normalization is now deterministic across both in-memory mutation and later
  sheet rename flows. GridGrind now normalizes POI-cloned calculated-column body formulas back to
  its metadata-owned table shape, rewrites copied totals-row structured references from POI's
  transient clone names such as `Table2` to the final copied table names, and keeps copied sheets
  rename-safe instead of failing with structured-reference parser crashes or dead transient table
  names.
- Drawing media and container signature-line authoring are now reliable across the packaged fat
  JAR and Docker image. GridGrind now rebuilds POI's private workbook picture catalog from actual
  `/xl/media/*` parts before picture-backed mutations so signature-line preview images cannot
  collide with later picture or embedded-object preview creation, and the Docker image now ships
  fontconfig plus DejaVu fonts so `SET_SIGNATURE_LINE` works in headless container runs. Docker
  smoke now verifies signature-line authoring directly.
- The largest remaining protocol and Jazzer god-files are now split at stable seams instead of
  continuing to accumulate unrelated responsibilities. Nested/plain protocol field-group catalog
  descriptors now live in dedicated support types, Jazzer workbook-operation generation now
  separates orchestration from value-payload factories, and build-failing architecture audits now
  keep contract/foundation boundaries, formula-write centralization, POI private-access
  centralization, and the new file-size ceilings from silently regressing.
- Jazzer verification now keeps its coverage contract aligned with the intended deterministic
  support subset after the operation-generator split. The extracted
  `OperationSequenceValueFactory` now shares `OperationSequenceModel`'s coverage scope, and new
  selector-sweep seam tests directly exercise the extracted value factory plus workflow/command
  cleanup paths so the refactor stays verified without distorting the Jazzer coverage gate.

## [0.53.0] - 2026-04-22

### Fixed

- Public Markdown docs are now organized as stable entry-point docs plus focused long-form
  references. `docs/OPERATIONS.md` and `docs/QUICK_REFERENCE.md` are now kept intentionally small,
  the quick-start flow now teaches artifact-native `--print-example BUDGET`, the error reference
  now documents the live `CELL_NOT_FOUND` path truthfully, the Java 26 workstation guide now
  reflects the current official OpenJDK 26.0.1 download surface, and docs tests now fail on broken
  local Markdown links or if the entry-point reference docs bloat back into god-files.
- Worksheet comments now survive structural column edits deterministically even across the Apache
  POI 5.5.1 collision case where moved comments can reopen with duplicate persisted comment refs
  and the wrong visible note after `INSERT_COLUMNS`, `DELETE_COLUMNS`, or `SHIFT_COLUMNS`.
  GridGrind now rewrites the affected comment/VML state authoritatively after those edits, the
  discovered `.xlsx` round-trip finding is promoted as a committed success seed, and the Jazzer
  round-trip verifier now fails saved workbooks whose persisted comment refs are not unique.

## [0.52.0] - 2026-04-22

### Fixed

- Chart readback now resolves `XSSFGraphicFrame -> chart relation -> XSSFChart` directly instead
  of scanning every chart relation and trusting POI's optional `chart.getGraphicFrame()`
  backpointer. Mixed live-chart plus frameless-chart states now degrade cleanly instead of
  crashing factual drawing/chart inspection, explicit graphic-frame snapshotting no longer loses
  formula-resolution context when POI leaves the chart backpointer unset, the reproduced
  frameless-chart relation crash is now promoted as a committed success seed, and multi-chart
  copied sheets are now covered by focused regressions so orphaned chart-frame associations
  cannot drift back in quietly after the `0.51.0` release.
- Excel-facing sheet layout limits are now consistent across the contract, engine, CLI catalog,
  and docs. Row height now enforces Excel's real `409.0`-point ceiling instead of POI's wider
  twip storage envelope, `SET_SHEET_PRESENTATION.sheetDefaults` now validates the same default
  row/column limits as explicit sizing commands, `SET_SHEET_ZOOM` is now registered in the public
  limitations registry, and malformed drawing packages with empty embedded-object bytes now
  degrade into truthful zero-byte readback instead of crashing workbook inspection or Jazzer
  round-trip verification.
- The release protocol now covers the real bootstrap case where the primary checkout holds the
  unpublished release payload but release verification must run from a clean worktree. The
  documented flow now requires an explicit bootstrap branch or exported patch before applying that
  payload in the release worktree, so a dirty or problematic primary checkout no longer forces
  operators into an undocumented release path.

## [0.51.0] - 2026-04-22

### Fixed

- Sheet copy now repairs embedded-object sheet relationships that Apache POI XSSF `cloneSheet`
  leaves behind for OLE package parts, so copied embedded objects survive in-memory inspection,
  save/reopen, and the reproduced Jazzer `COPY_SHEET` round-trip regression. The failing fuzz case
  is now promoted as a committed success seed.
- The release closeout protocol is now executable instead of prose-only. GridGrind now ships
  `scripts/verify-release-primary-checkout.sh`, a dedicated shell regression for it, and updated
  release docs/check wiring so releasing from a disposable worktree cannot quietly leave the
  primary checkout behind `origin/main` with stale version-bearing files and misleading overlays.
- The supported Jazzer wrapper surface is now part of the tracked repository instead of hidden
  behind ignored local `jazzer/bin/` files. Clean release worktrees now carry the same
  `jazzer/bin/*` operator scripts as the primary checkout, and a new regression fails if the
  documented Jazzer wrapper surface drifts from what a clean checkout actually ships.

## [0.50.0] - 2026-04-22

### Fixed

- GridGrind now exposes Apache POI/XSSF workbook custom-XML mapping workflows as first-class
  public contract operations. `IMPORT_CUSTOM_XML_MAPPING`, `GET_CUSTOM_XML_MAPPINGS`, and
  `EXPORT_CUSTOM_XML_MAPPING` are now wired end to end across the contract, engine, executor,
  shipped examples, and public docs, with source-backed XML import support and focused regression
  coverage.
- Dedicated array-formula workflows are now documented truthfully across the public surface.
  README, reference docs, and the POI/XSSF capability inventory now distinguish scalar
  `FORMULA` cell writes from `SET_ARRAY_FORMULA` / `CLEAR_ARRAY_FORMULA` / `GET_ARRAY_FORMULAS`
  instead of incorrectly describing array formulas as wholly unsupported.
- The Apache POI/XSSF public docs now match the audited Apache POI 5.5.1 surface more
  truthfully. `docs/POI_EXCEL_CAPABILITY_INVENTORY.md` no longer overstates upstream support for
  sparklines or threaded comments, now splits custom XML mappings from slicers, and now lists
  previously omitted non-productized families such as signature lines and broader XDDF chart
  authoring. `docs/LIMITATIONS.md` now states that row/column worksheet bounds are already
  enforced on relevant request paths, and it no longer implies that Apache POI XSSF can only
  handle plain `.xlsx` files. Contract-side doc and POI-runtime audit tests now fail the build if
  these claims drift again or if the upstream jar surface changes underneath the published
  inventory.
- Sheet copy now rides POI XSSF's native clone path for the drawing family and then reapplies
  GridGrind-owned repair passes for workbook-core details such as formulas, raw data validations,
  conditional formatting, tables, sheet-owned autofilters, local names, comments, print layout,
  and protection metadata. Copied sheets now preserve supported pictures and charts instead of
  silently dropping them.
- Multi-plot combo-chart authoring is now verified directly in the engine test suite, and the
  public POI/XSSF capability inventory and operations docs now describe combo charts as a shipped
  supported capability instead of an undocumented gap.
- Signature-line drawing metadata is now exposed as a first-class `.xlsx` surface. GridGrind now
  ships `SET_SIGNATURE_LINE`, factual `SIGNATURE_LINE` drawing-object readback, authored anchor
  replacement and delete support for named signature lines, a generated signature-line example,
  and public docs that describe the real XSSF picture-format surface (`GIF`, `TIFF`, `EPS`,
  `BMP`, and `WPG` included) instead of the old narrower list.
- Release-surface shell verification is now resilient on this macOS baseline too. The packaged
  CLI/publication verifiers and their regression scripts now use repo-local disposable scratch
  under `tmp/` instead of brittle `/var/folders` temp allocation, and the release merge-handoff
  plus candidate-tag regression harnesses now fake the remote Git fetch surface directly instead
  of depending on flaky local clone/push transport behavior.
- `--print-protocol-catalog --operation` now exposes the nested and plain type groups that
  operators actually need for black-box request authoring. Type-group lookups such as
  `nestedTypes:cellInputTypes` and `plainTypes:chartInputType` now work directly instead of
  forcing callers to download and parse the full catalog just to discover cell, chart, or source
  payload shapes.
- Formula-backed chart references now evaluate authoritatively during chart authoring and
  readback instead of trusting stale worksheet or OOXML chart caches. Reference-backed series and
  formula titles now round-trip deterministically through save/reopen and the reproduced Jazzer
  crash is promoted as a committed success regression seed.
- Root `./check.sh` stall diagnostics now own diagnostic subprocess lifecycles end to end. The
  timeout path was extracted into a dedicated process-support helper, `capture_with_timeout` now
  escalates from `TERM` to `KILL` across the full captured process tree, and a shell regression
  now proves that TERM-ignoring parent/child trees do not leave orphaned `jcmd`-style processes
  behind.

## [0.49.0] - 2026-04-21

### Changed

- GridGrind now publishes a contract-owned intent-discovery layer on top of the exact protocol:
  `--print-task-catalog`, `--print-task-plan <id>`, `--print-goal-plan "<goal>"`, and
  `--doctor-request`. The packaged JAR and Docker artifact verifiers now black-box these task,
  planning, and diagnostics surfaces too, so the intent layer cannot drift silently from the core
  protocol contract.

### Fixed

- The thin CLI and packaged artifact help surface now behave truthfully for black-box operators:
  invoking GridGrind with no arguments and empty stdin prints help instead of falling through into
  a JSON parse failure, and `--print-protocol-catalog --operation <id>` now rejects ambiguous raw
  ids explicitly while teaching the canonical `<group>:<id>` lookup form.
- Goal-plan normalization now handles common plural stems more truthfully for freeform task
  discovery, including `quizzes -> quiz` and `sizes -> size`, instead of degrading ranked task
  matches with avoidable stemming drift.
- Chart readback now prefers live worksheet values over stale embedded chart caches for
  reference-backed series, so packaged artifact smoke and operator readback reflect the workbook's
  current sheet state instead of outdated OOXML cache text.
- Failure `context.stage` is now truthful and round-trippable for `RESOLVE_INPUTS`,
  `CALCULATION_PREFLIGHT`, and `CALCULATION_EXECUTION`; the calculation stage discriminator is now
  owned by concrete context types instead of a mismatched dynamic field.
- Synthetic CLI failure journals no longer invent `"unknown-plan"` or `"UNKNOWN"` source and
  persistence types before a request has parsed. Pre-parse failures now omit unavailable plan and
  transport facts instead of serializing misleading sentinel values.
- Step-target parsing now reports disallowed selector types truthfully, including ambiguous
  selector-family overlaps such as `BY_NAME`, instead of swallowing earlier parse failures and
  blaming the last attempted selector shape.
- `WorkbookPlan.steps()` now stays immutable after construction, so callers can no longer mutate a
  parsed or authored plan in place and bypass the duplicate-`stepId` invariant.
- Request-stream transport is now bounded at 16 MiB. Oversized JSON requests fail early with a
  product-owned `INVALID_REQUEST` message, and the streamed request/response/catalog write APIs no
  longer build whole-payload byte arrays before writing to caller-owned output streams.
- The machine-readable catalog and CLI help now publish step-target selector families directly.
  `mutationActionTypes`, `assertionTypes`, and `inspectionQueryTypes` now expose
  `targetSelectors` and `targetSelectorRule`, table and pivot mutations now declare only
  `BY_NAME_ON_SHEET`, and ambiguous shared selector ids such as `BY_NAME` are rejected explicitly
  instead of being guessed by parser order.
- `WorkbookStepJsonDeserializer` now uses Jackson 3's non-deprecated string-node APIs, and
  print-layout margin reads and writes now use Apache POI's `PageMargin` enum surface instead of
  the deprecated short-constant overloads. `ExcelStreamingWorkbookWriter` no longer carries a
  deprecated no-op `Cell.setCellType(FORMULA)` call after `setCellFormula(...)`, and its
  `close()` path now relies on `SXSSFWorkbook.close()`, which already disposes SXSSF temp files in
  Apache POI 5.5.1.
- `GridGrindJson` no longer emits `AlmostJavadoc` warnings from package-private parser-message
  helpers. The production JSON wording helpers now use real Javadoc comments so static analysis
  output stays signal-rich.
- Stale promoted Jazzer metadata for invalid protocol-request seeds is now refreshed against the
  current selector contract. The refresh test fixture now records the modern
  `EXPECTED_INVALID`/`INVALID_REQUEST_SHAPE` replay surface instead of preserving pre-hard-break
  authored source facts after parsing has already failed.
- Encrypted `.xlsx` package handling now models unencrypted versus password-protected open state
  explicitly instead of relying on nullable password sentinels, and POI relation-removal failures
  no longer mask fatal JVM errors as ordinary state exceptions.
- Exact-cell read validation is now single-sourced end to end. `GET_CELLS` address lists now share
  one canonical duplicate, bounds, and shape validator across the contract and engine while still
  preserving the indexed request messages callers already rely on.
- The Java workbook engine now exposes formula evaluation, capability inspection, cache clearing,
  and recalc-on-open toggles through a dedicated `workbook.formulas()` surface instead of crowding
  those operations onto the root workbook wrapper, and distinct gradient fills now keep their own
  OOXML entries even when Apache POI's public equality checks would otherwise alias different
  gradient geometries together.
- Executor and engine internals are now split into narrower helpers for execution validation,
  request path facts, calculation workflows, response shaping, sheet annotations, prepared chart
  models, chart source and snapshot handling, drawing removal and binary cleanup, drawing-chart
  flows, and drawing anchor state. The public contract stays the same, but the request executor
  and workbook engine are less monolithic and easier to evolve safely.
- Canonical operation, assertion, and inspection ids, execution-mode limits, CLI help labels, and
  shared discovery/error wording now derive from contract-owned structured metadata instead of
  parallel string copies. The CLI help surface is now a typed catalog section tree, execution-mode
  validation messages come from one shared metadata owner, and the build now fails if public
  docs/help/catalog/example surfaces mention an unregistered canonical mutation, assertion, or
  inspection id.

## [0.48.0] - 2026-04-19

### Changed

- GridGrind no longer ships a monolithic `protocol` module. The public contract, metadata
  registry, and JSON codecs now live in `contract`, while `executor` is the only execution bridge
  into the workbook engine. This also puts the Java-first contract replacement program into formal
  hard-break mode.
- The nested Jazzer build now consumes the `executor` plus `contract` split instead of the deleted
  monolithic `protocol` module, and the `contract` module's 100 % coverage gate now includes the
  downstream `executor` and `cli` consumer tests that exercise the canonical public contract after
  the split.
- The public contract is now selector-first end to end. Top-level coordinate and sheet-name
  fields have been removed from request operations and reads in favor of canonical selector
  payloads, and the public README, operations reference, quick reference, shipped examples, and
  release gates now reject the deleted pre-selector vocabulary.
- The public request model is now step-only end to end. The deleted `operations[]` plus `reads[]`
  split has been replaced by ordered `steps[]`, step envelopes no longer carry a redundant outer
  discriminator, and public request/response docs, Docker smoke, replay fixtures, and contract
  surface checks now reject the deleted `selector` and `requestId` vocabulary alongside the old
  arrays.
- GridGrind now ships first-class assertion steps in the public contract. Ordered plans can verify
  workbook state with canonical `ASSERTION` steps, success responses return ordered
  `assertions[]`, and public docs, examples, CLI help, and artifact-surface verification now
  describe the mutate-then-verify model instead of the older mutation-plus-read-only framing.
- GridGrind responses now always carry a structured execution `journal`, and the top-level request
  execution policy is now the canonical `execution` envelope with nested `mode` and `journal`
  configuration instead of the older top-level `executionMode` sidecar. `VERBOSE` journal mode is
  now a public contract feature that streams fine-grained live execution events to CLI stderr
  while preserving the same events in the response payload.
- GridGrind no longer exposes formula evaluation, cache clearing, or recalc-on-open as mutation
  step types. Those deleted actions are replaced with the canonical top-level
  `execution.calculation` policy, whose shipped strategies are `EVALUATE_ALL`,
  `EVALUATE_TARGETS`, `CLEAR_CACHES_ONLY`, and `DO_NOT_CALCULATE` plus optional
  `markRecalculateOnOpen=true`.
- GridGrind mutation payloads are now source-backed end to end. Text-bearing authored fields use
  canonical `INLINE`, `UTF8_FILE`, or `STANDARD_INPUT` sources, binary-bearing authored fields use
  canonical `INLINE_BASE64`, `FILE`, or `STANDARD_INPUT` sources, and the structured execution
  journal now records authored input loading under `journal.inputResolution` before workbook open.
- GridGrind now ships the `authoring-java` module as a first-class fluent Java surface over the
  canonical contract. Java callers can build selector-first mutation, inspection, assertion, and
  execution-policy plans without hand-writing JSON. The authored target records now reject null
  selectors immediately instead of carrying invalid state until later execution, and the shipped
  `examples/java-authoring-workflow.java` example is compile-verified against that published API.
- CLI help, built-in examples, and the machine-readable catalog are now contract-owned surfaces.
  The thin CLI transport renders help from the `contract` module, `--print-example <id>` emits one
  generated built-in example request, `--print-protocol-catalog` publishes `cliSurface` plus
  `shippedExamples`, and the committed `examples/*.json` fixtures are regenerated from that same
  registry instead of being hand-maintained JSON.
- The root `check` gate now includes explicit-import verification for handwritten production Java
  and Kotlin sources, so wildcard imports fail the canonical build instead of relying on reviewer
  cleanup.
- Release and fuzz hardening now treat packaged artifacts as the authoritative contract surface.
  CLI contract verification reads structured help plus shipped-example lines from the built JAR or
  Docker image, Docker smoke rebuilds the CLI fat JAR before packaging so it cannot validate stale
  output, and promoted Jazzer metadata refresh rewrites dead source pointers to the live promoted
  inputs when regenerated examples retire old fixture paths.
- Named-range deletion now accepts only exact scoped selectors in the canonical Java contract.
  Workbook-scoped and sheet-scoped named ranges remain the same JSON wire types, but the hard-
  break Java surface no longer allows broad named-range selector families where authoritative
  deletion semantics require an exact scope.
- The committed `protocol-workflow` Jazzer corpus now uses neutral `workflow_case_##` seed names
  instead of semantic blob names. Those inputs are opaque generator bytes, so the authoritative
  decoded behavior now lives only in refreshed replay metadata and replay text.
- Jazzer replay metadata and replay text for protocol request/workflow harnesses now record
  assertion counts and assertion kinds explicitly, so assertion-bearing opaque workflow seeds no
  longer look like unexplained response-kind flips during regression replay.
- The release protocol now treats open Dependabot PRs as first-class release hygiene. Release-time
  pre-flight now requires explicitly identifying open Dependabot work, and after the public
  release is verified each Dependabot PR must be merged, closed, or consciously kept open with a
  stated reason; stale automation branches are no longer acceptable release leftovers.
- Workbook-protection readbacks, docs, and parity checks now use the same canonical field name
  `revisionsLocked` end to end instead of mixing singular and plural variants across write and
  read surfaces.
- Formula-backed chart titles are now authored with an explicit OOXML string cache, so numeric
  title cells survive `.xlsx` save/reopen exactly instead of drifting to stale cached values
  during chart round-trips.

### Fixed

- Chart reads now resolve formula-backed title `cachedText` from the referenced cell when OOXML
  omits a cached string, so authored chart titles no longer come back blank while the same chart's
  formula still points at a real text cell.
- Existing bar, line, and pie charts now treat `Title.None` series updates idempotently, so
  removing a series title no longer falls into Apache POI `unsetTx()` crashes when the underlying
  chart XML never carried a `<c:tx>` node.
- The committed Jazzer regression floor now includes the `fractional_integer_field.json` invalid-
  request seed with matching promotion metadata and replay text, so the nested fuzz build no
  longer drifts when integer-field shape validation expands.
- The nested Jazzer replay floor now refreshes both selector-first protocol-request fixtures and
  neutralized protocol-workflow binary metadata against the current replay engine, so contract
  topology changes no longer leave stale replay expectations behind.
- The Docker smoke gate now verifies low-memory streaming in the same two-step shape the product
  actually supports: `STREAMING_WRITE` authoring first, then summary-only `EVENT_READ` readback
  against the materialized workbook.
- Table-aware exact-cell selectors now execute truthfully end to end. Selector-first table-key
  targeting resolves through the canonical executor for mutations, inspections, and assertions,
  source-backed row-key values now load before execution just like source-backed mutation payloads,
  duplicate-key matches fail instead of guessing, and zero-match inspections no longer collapse
  into lower-level table lookup failures.
- The Docker smoke and publication-surface regression gates now author `APPEND_ROW` text values
  with canonical source-backed payloads.
- Streaming-write calculation failures now surface as structured `CALCULATION_EXECUTION` contract
  failures before persistence instead of collapsing into generic request-level runtime handling.
- Drawing-object picture reads now report factual raster dimensions when the image format exposes
  them, and expanded row grouping no longer destroys pre-existing manual hidden-row state during
  group or ungroup operations.
- Protocol JSON parsing now rejects floating-point JSON numbers for integer contract fields instead
  of silently truncating them during deserialization.
- Selector and operation validation now preserve indexed null and invalid-entry diagnostics instead
  of letting early collection copying collapse those failures into opaque bare `NullPointerException`
  paths.
- The workbook-core inspection pipeline now uses `stepId` internally as well as on the public
  contract, so command/result plumbing, executor conversion, parity helpers, and inspection tests
  no longer translate back through the deleted `requestId` terminology.
- The nested Jazzer workflow generator, invariants, telemetry, and replay-safe support layer now
  exercise assertion steps as first-class workflow elements instead of fuzzing only mutations plus
  inspections, and the shipped `examples/assertion-request.json` fixture is regression-tested as a
  public mutate-then-verify request.
- Presence-style assertions now own selector-count semantics authoritatively: exact named-range
  and chart misses are evaluated as zero observed entities for
  `EXPECT_NAMED_RANGE_PRESENT` / `EXPECT_NAMED_RANGE_ABSENT` and
  `EXPECT_CHART_PRESENT` / `EXPECT_CHART_ABSENT` instead of leaking lower-level not-found
  failures, and the executor regression suite now locks that behavior in.
- Every response surface now exposes execution journaling consistently: success and failure payloads
  include structured phase and step telemetry, the shipped assertion example demonstrates
  `execution.journal.level=VERBOSE`, and Docker smoke now black-boxes both response-journal
  emission and live stderr event streaming from the packaged artifact.
- Source-backed authored input failures now surface as explicit `INPUT_SOURCE_NOT_FOUND`,
  `INPUT_SOURCE_UNAVAILABLE`, or `INPUT_SOURCE_IO_ERROR` contract failures instead of collapsing
  into generic request or filesystem errors, and the shipped
  `examples/source-backed-input-request.json` fixture plus promoted Jazzer replay floor now lock
  file-backed text, file-backed formulas, and file-backed binary payloads into regression coverage.
- Engine selection and read payload helpers no longer perform wasteful double-freeze list copies
  when validating already-immutable `List.copyOf(...)` inputs.
- The build and developer docs now spell out the upstream Jackson 3 rule correctly: Jackson 3
  databind intentionally still uses the `com.fasterxml.jackson.annotation` namespace, so GridGrind
  now guards that fact explicitly instead of making it look like accidental dependency drift.

## [0.47.0] - 2026-04-16

### Changed

- Public streaming-write docs, CLI help, the protocol catalog, and runtime validation messages now
  use the single canonical operation name `FORCE_FORMULA_RECALCULATION_ON_OPEN` consistently.
- Agent-facing catalog summaries now describe the actual response shape for `GET_SHEET_LAYOUT`,
  `GET_FORMULA_SURFACE`, `GET_NAMED_RANGE_SURFACE`, and the workbook-health analysis reads,
  including checked-count fields, `layout.presentation`, grouped formula summaries, and flat
  aggregated findings.
- Core-owned contract text now drives the thin downstream CLI help, protocol catalog discovery
  summaries, and execution-mode validation messages for the public low-memory and workbook-health
  surfaces, so operation-name and limitation wording no longer lives in multiple hand-maintained
  string copies.
- Release verification now treats `--help` and `--print-protocol-catalog` as first-class public
  artifacts. The build JAR, local Docker smoke image, and published GHCR tags are black-box
  checked for canonical streaming-write wording, the hard `LAMBDA`/`LET` boundary, and the richer
  analysis/layout catalog summaries before a release is considered healthy.

### Fixed

- Gradient-fill authoring now rejects mixed geometry models up front instead of serializing
  impossible `.xlsx` styles. `LINEAR` gradients accept `degree`, `PATH` gradients accept
  `left/right/top/bottom`, and valid gradient-plus-protection styles now survive `.xlsx`
  round-trips cleanly.
- Formula docs and CLI help now state the current hard limitation truthfully: authored
  array-formula braces are rejected as `INVALID_FORMULA`, `LAMBDA` and `LET` are currently
  rejected as `INVALID_FORMULA` because Apache POI cannot parse them, and loaded formulas that POI
  parses but cannot evaluate surface as `UNSUPPORTED_FORMULA`.
- Protocol JSON parsing no longer carries a special-case alias hint path for retired operation
  discriminator spellings; invalid discriminators now fail uniformly as unknown type values.
- The Apache POI XSSF capability inventory now cites the real comment implementation files instead
  of a stale nonexistent comment-support evidence path.

## [0.46.0] - 2026-04-15

### Changed

- The release protocol now requires an explicit post-merge `main` CI handoff before tagging, and
  `scripts/verify-release-merge-handoff.sh` plus its shell regression test now enforce that the
  checked-out release commit matches `origin/main` and already has green `Check` and `Docker smoke`
  runs before any public tag is created.
- The machine-readable protocol catalog and CLI help now describe the full
  `ANALYZE_WORKBOOK_FINDINGS` aggregate correctly, including pivot-table health, and the public
  workbook-health docs now point directly at the shipped batch-analysis request examples.

### Fixed

- The public Apache POI XSSF capability inventory no longer collapses distinct read and analysis
  operations into shared rows. It now names the individual formula, hyperlink, named-range, sheet
  schema, and aggregate workbook-finding surfaces explicitly, and it declares array formulas and
  sparklines as `NOT_EXPOSED` instead of leaving them implicit.
- Public formula docs now spell out the real request-contract boundary: authored formulas are
  scalar only, array-formula braces are rejected as `INVALID_FORMULA`, newer constructs such as
  `LAMBDA`/`LET` may fail during parse when Apache POI cannot read them, and loaded formulas that
  POI parses but cannot evaluate surface as `UNSUPPORTED_FORMULA`.
- The committed `examples/*.json` request fixtures are now regression-tested for JSON/schema
  validity so public examples cannot drift silently from the shipped protocol.

## [0.45.0] - 2026-04-14

### Changed

- Release and container workflow-dispatch reruns now refuse to publish unless the checked-out tag
  still matches `gradle.properties`, resolves to the exact remote tag commit, remains reachable
  from `main`, and already has successful `Check` plus `Docker smoke` CI runs on that commit.
- Published GHCR release images now rebuild from a digest-pinned Azul Java 26 base image and emit
  explicit OCI provenance plus SBOM attestations.
- The root `coverage` and `jacocoAggregatedReport` tasks now discover JaCoCo-enabled Java
  subprojects dynamically and aggregate every module `build/jacoco/*.exec` file instead of
  hardcoding today's module list.
- The `protocol` module no longer leaks `jackson-databind` onto consumer compile classpaths; the
  runtime dependency remains available where needed without widening the published Gradle API.

### Fixed

- The container publication workflow no longer overrides the image's SPDX license label with a
  stale MIT-only value; published OCI metadata now matches the shipped notices and licenses.
- `./check.sh` Stage 4 now regression-tests the release-candidate tag gate and publication
  workflow contract, so pinned-base-image, attestation, and publication-policy drift fails
  locally before a public release run.

## [0.44.0] - 2026-04-14

### Changed

- JaCoCo branch-coverage enforcement raised toward 100 % for the `engine` module: branch paths
  previously unreachable by the test suite are now exercised across seven controller and support
  classes (details below). Coverage targets are asserted by the Gradle `coverage` task and must
  remain green before any release.
- The remaining advanced-XSSF branch gaps are now closed across drawing, pivot-table,
  package-security, autofilter, print-layout, data-validation, sheet-state, event-read,
  sheet-copy, conditional-formatting, formula-rename, and formula-exception flows. Helpers that
  were only guarding unmaterializable POI or XmlBeans null states were simplified to match the
  real runtime contract, while new tests now exercise pivot-cache registration cleanup, embedded
  preview lookup, sparse event-reader workbook metadata, and package-signing or encryption edge
  handling.
- `ExcelPrintLayoutController` branch coverage extended: `shouldUnsetPageSetupOrientation` now
  covers the `LANDSCAPE` false-return path in addition to `PORTRAIT`; `isEmptyPageSetup` now
  covers all eleven individual attribute-present/default combinations (paperSize, draft,
  blackAndWhite, copies, useFirstPageNumber, firstPageNumber) for both the true and false returns.
- `ExcelConditionalFormattingController` health-check branch coverage extended: the
  `CellValueRule`-with-two-valid-formulas path (both `isBrokenFormula` calls return false, no
  health finding emitted) is now exercised by a dedicated test so the short-circuit `&&` chain
  cannot regress silently.
- `ExcelWorkbook` persistence-options branch coverage extended: the
  `persistenceOptions != null && persistenceOptions.isEmpty()` fast-path (treat as plain save)
  is now covered by a test that passes `new ExcelOoxmlPersistenceOptions(null, null)` to
  `save(Path, ExcelOoxmlPersistenceOptions)`.
- `ExcelFormulaSheetRenameSupport` external-workbook branch coverage extended: the
  `getExternalWorkbookNumber() >= 1` guard (skip rename for cross-workbook references) is now
  covered via a test that links an external workbook, references it from a formula, and asserts
  the returned formula still names the external book rather than the renamed local sheet.
- `FormulaExceptions.isKnownBuiltinFunction` branch coverage extended: the
  `AnalysisToolPak.getSupportedFunctionNames()` and `AnalysisToolPak.getNotSupportedFunctionNames()`
  OR-chain branches (conditions 4 and 5) are now covered by tests that locate functions present
  exclusively in each ATP set — absent from `FunctionMetadataRegistry` and both
  `FunctionEval` catalogs — so each OR branch can be the first-true entry.
- `ExcelSheetCopyController` validation-formula retargeting branch coverage extended: edge cases
  now tested include formulas shorter than two characters, quoted-list literals that start but do
  not end with `"`, absent `formula1`, blank `formula1`, absent `formula2`, and blank `formula2`,
  exercising all guard branches in `isQuotedListLiteral` and the `isSetFormula`/blank checks.
- `ExcelEventWorkbookReader` branch coverage extended: the `sizeOfWorkbookViewArray() == 0` path
  in `activeSheetIndex` is now covered; additionally an integration test exercises blank cell-ref
  (`r=""`) handling and a `<col>` element with no `max` attribute in the sheet XML parse path.
- README, the public capability inventory, the internal XSSF parity ledger, and the Jazzer
  coverage guide now agree on the shipped verification surface for drawings, charts, pivots,
  conditional formatting, low-memory event reads, and OOXML package security. The Jazzer guide no
  longer lists already covered chart, picture, pivot, or conditional-format families as missing
  fuzz surface.

### Fixed

- Docker smoke now builds through `docker buildx build --load` under an anonymous `DOCKER_CONFIG`
  while preserving the active local Docker engine endpoint, so local and CI verification no longer
  depend on personal Docker login state or Docker's deprecated legacy builder path.
- Docker smoke now runs mounted-path container commands as the caller's UID:GID, so generated
  response files and saved workbooks stay operator-owned instead of leaving root-owned artifacts
  behind on Linux CI or other Unix-like hosts.
- `scripts/verify-container-publication.sh`, the Docker developer docs, and the release protocol
  now codify anonymous public-container verification while preserving the active Docker engine
  endpoint, keeping release verification aligned with the real workstation contract.
- `NOTICE` now covers all bundled runtime dependencies: Apache Santuario xmlsec (Apache 2.0),
  Bouncy Castle bcpkix/bcprov/bcutil (MIT variant), and SLF4J API (MIT). These were added as
  runtime dependencies in 0.43.0 but omitted from the attribution file.
- `--license` output now includes the `NOTICE` attribution file between the GridGrind MIT license
  and the third-party license texts, satisfying Apache License 2.0 §4(d) for runtime distributions.
- `PATENTS.md` dependency table now lists Apache Santuario xmlsec, Bouncy Castle, and SLF4J API.
- Fat JAR (`META-INF/`) now includes `PATENTS.md` alongside `NOTICE` and the license files.
- Docker image now ships `PATENTS.md` in `/usr/share/doc/gridgrind/` for completeness.
- `org.opencontainers.image.licenses` OCI label updated from `MIT` to the full SPDX expression
  `MIT AND Apache-2.0 AND BSD-3-Clause` to accurately reflect the bundled dependency licenses.
- `scripts/verify-container-publication.sh` now verifies the exact two-line `--version` product
  header that the shipped container exposes (`GridGrind <version>` plus the product
  description), and `./check.sh` now includes a dedicated shell regression for that verifier so
  future release workflows cannot fail on a stale output assumption.
- The packaged CLI and Docker image now include the Bouncy Castle runtime dependencies required by
  OOXML package-security inspection, so `source.type=EXISTING` opens and low-memory
  `STREAMING_WRITE` readback no longer fail at runtime with `NoClassDefFoundError` on the
  existing-workbook path.
- The packaged CLI now also bundles an explicit SLF4J provider for package-security dependencies,
  and the Docker smoke gate rejects unexpected stderr on request execution, so successful
  existing-workbook and low-memory request flows stay clean for agent consumers.
- The Jazzer runtime now binds the same SLF4J provider explicitly and asserts it in support
  tests, so regression replay and round-trip harnesses no longer emit fallback logger warnings
  on stderr while they pass.
- The machine-readable protocol catalog now correctly marks `SET_SHEET_PROTECTION.password` and
  `SET_AUTOFILTER.criteria` or `sortState` as optional to match the shipped runtime contract.
- Public request-shape docs and catalog summaries now clarify sparse append-edge row or column
  inserts: valid append-position inserts do not materialize a new physical tail row or column
  until content or explicit metadata exists there.

## [0.43.0] - 2026-04-13

### Added

- `--license` flag: prints the GridGrind MIT license followed by third-party dependency license
  notices (Apache 2.0, BSD-3-Clause). License texts are bundled in the JAR under
  `licenses/` and read at runtime.
- Added `SET_SHEET_PRESENTATION` plus `GET_SHEET_LAYOUT.presentation` for sheet display flags,
  right-to-left layout, tab color, outline-summary placement, default row or column sizing, and
  ignored-error suppression.
- Added `printGridlines` to print-setup authoring and factual `GET_PRINT_LAYOUT` readback.

### Changed

- `--version` now prints the same two-line product header as the opening of `--help`
  (`GridGrind <version>` on the first line, the product description on the second), providing
  a single source of truth for product identity output via `productHeader()`.
- `--help` banner changed from `GridGrind CLI <version>` to `GridGrind <version>` to match
  the product name.
- Promoted the Apache POI XSSF capability inventory into a public reference at
  [docs/POI_EXCEL_CAPABILITY_INVENTORY.md](./docs/POI_EXCEL_CAPABILITY_INVENTORY.md) and rewrote
  it around the shipped `.xlsx` contract instead of internal parity-phase language.
- README now links directly to the public capability inventory alongside the operations,
  quick-reference, and limitations docs.
- Public docs, quick-reference snippets, and runnable structural-layout examples now describe the
  shipped sheet-presentation surface and print-gridline output explicitly, and the README example
  guide no longer uses internal phase labels.

### Fixed

- The parity-doc regression test now verifies the public capability inventory path, release
  version, public-facing wording, and the absence of the retired internal parity-planning files.
- The POI capability inventory no longer underreports the shipped sheet-view and print-layout
  surface; it now reflects the public contract accurately.
- Sheet-presentation and print-setup model validation no longer carry dead compatibility overloads
  or unreachable null-element branches, so the Java contract stays smaller and the regression floor
  asserts only live behavior.

## [0.42.0] - 2026-04-13

### Added

- Added OOXML package-security support to the public `.xlsx` contract:
  `source.security.password` for encrypted existing sources, `persistence.security.encryption`,
  `persistence.security.signature`, and factual `GET_PACKAGE_SECURITY` readback for package
  encryption and package-signature state.
- Added `examples/package-security-create-request.json`
  and [examples/package-security-inspect-request.json](./examples/package-security-inspect-request.json),
  a paired public example flow for encrypted workbook authoring followed by factual package-security
  inspection.
- Added a promoted Jazzer protocol-request seed for the package-security request surface so
  `source.security`, `persistence.security`, and `GET_PACKAGE_SECURITY` stay replay-verified in the
  committed regression floor.

### Changed

- Public docs, quick-reference snippets, README guidance now describe the shipped OOXML
  package-security surface explicitly instead of treating encryption and signing as absent.

### Fixed

- Legacy OLE2 `.xls` files are no longer misclassified as encrypted OOXML packages on the
  package-security open path; unsupported legacy workbooks now fail honestly as non-`.xlsx`
  inputs instead of incorrectly demanding `source.security.password`.
- The parity ledger now verifies actual encrypted-open, encrypted-save, signed-read,
  signed-authoring, invalid-password, and invalid-signature behavior against the committed corpus
  instead of keeping legacy "gap" probes after the runtime support landed.
- The invalid-signature parity corpus now tampers signed workbooks through a separate output path
  instead of rewriting the signed package in place, so the corpus materializes as genuinely
  `INVALID` rather than as a broken ZIP stream.

## [0.41.0] - 2026-04-13

### Added

- Added top-level `executionMode` request support for low-memory `.xlsx` workflows:
  `readMode: EVENT_READ` for summary-only event-model reads and `writeMode: STREAMING_WRITE` for
  append-oriented SXSSF authoring on `NEW` workbooks.
- Added [examples/large-file-modes-request.json](./examples/large-file-modes-request.json), a
  runnable example covering `executionMode`, `STREAMING_WRITE`, and summary-only `EVENT_READ`
  readback.

### Changed

- Public docs, CLI help, the limitations registry now describe the shipped low-memory execution
  contract explicitly instead of treating event reads and streaming writes as absent.

### Fixed

- `ExecutionModeInput.isDefault()` is now JSON-ignored, so request round-trips no longer leak a
  stray `executionMode.default` field.
- The streaming writer now owns row advancement explicitly and disposes SXSSF temp files on close
  instead of relying on incidental POI state.
- Parity probes now verify actual low-memory request behavior against the large-sheet
  corpus instead of only checking for missing catalog placeholders.
- Jazzer replay support now resolves promoted large-file example inputs correctly from both the
  `jazzer` module root and the repository root, so the committed low-memory example stays
  replay-verified under the real Gradle execution layout.

## [0.40.0] - 2026-04-13

### Added

- Added limited XSSF pivot-table parity to the public `.xlsx` contract with `SET_PIVOT_TABLE`,
  `DELETE_PIVOT_TABLE`, `GET_PIVOT_TABLES`, and `ANALYZE_PIVOT_TABLE_HEALTH`.
- Added [examples/pivot-request.json](./examples/pivot-request.json), a runnable example covering
  range-backed, named-range-backed, and table-backed pivot authoring plus factual pivot readback
  and pivot-health analysis.
- Added the matching promoted Jazzer protocol-request seed so the public pivot example is
  replay-verified in the committed regression floor.

### Changed

- Public docs, quick-reference snippets, README guidance, now describe the shipped pivot-table
  surface explicitly, including supported source kinds, authored anchor rules, explicit
  unsupported readback, and pivot-health analysis.

### Fixed

- Pivot-table authoring and preservation now normalize named-range sources correctly, keep
  workbook-global pivot names and cache relations stable across save or reopen, and rebuild
  workbook pivot wiring after create or delete flows instead of depending on stale POI relation
  allocation state.
- Replacing a pivot table in place on the same sheet now cleans up orphaned cache-record parts and
  re-primes POI's pivot-part allocator from the workbook package, so repeated authoring no longer
  collides on stale `/xl/pivotCache/pivotCacheRecords*.xml` numbering after delete or reopen flows.
- Malformed or oversized pivot number-format identifiers now degrade into truthful readback instead
  of throwing during factual snapshotting or parity-oracle reporting.
- Jazzer request labeling, workflow invariants, workbook-shape checks, `.xlsx` round-trip
  verification, replay expectations, and promoted workflow metadata now model the current
  pivot-table and protocol-workflow contract directly instead of lagging behind the shipped
  behavior.

## [0.39.0] - 2026-04-12

### Added

- Added [examples/chart-request.json](./examples/chart-request.json), a runnable example covering
  supported `BAR` chart authoring, named-range-backed series binding, explicit chart-anchor
  replacement, factual `GET_CHARTS` readback, and matching chart inventory in
  `GET_DRAWING_OBJECTS`.
- Added the matching promoted Jazzer protocol-request seed plus deterministic Jazzer support
  coverage for chart authoring and chart readback, so the public example is replay-verified and
  the fuzz-support layer now exercises the chart contract directly.

### Changed

- Public docs, quick-reference snippets, and README guidance now document the shipped chart
  contract explicitly: `SET_CHART`, `GET_CHARTS`, supported simple `BAR` or `LINE` or `PIE`
  families, named-range-backed series formulas, and explicit `UNSUPPORTED` readback for
  unsupported plot families.
- `SET_SHAPE` and `SET_CHART` validation is now explicitly non-mutating: failed authored shape or
  chart requests leave existing drawing state untouched instead of leaking partial artifacts.

### Fixed

- Jazzer `.xlsx` round-trip verification now asserts chart and drawing-object preservation across
  reopen instead of treating charts as an untracked blind spot.
- Jazzer request labeling, workflow-shape validation, and workbook-shape invariants now model
  `SET_CHART`, `GET_CHARTS`, and chart-backed drawing inventory as first-class protocol surface.
- Failed `SET_SHAPE` preset validation and failed `SET_CHART` preflight no longer leave partial
  shapes, chart frames, or half-mutated existing charts behind.
- Chart factual reads now normalize blank stored OOXML titles to `NONE`, preserve sparse literal
  cache positions as empty-string gaps instead of aborting the read, and degrade broken chart
  relationships into truthful surviving drawing facts when a graphic frame remains.
- The chart controller now owns explicit POI translation and relation-removal seams, so
  chart-family enum mapping and chart-part deletion are regression-tested directly instead of
  hiding inside one large controller branch.

## [0.38.0] - 2026-04-12

### Added

- Added `examples/drawing-media-request.json`, a runnable
  example covering picture, shape, and embedded-object authoring, explicit drawing-anchor
  replacement, drawing payload extraction, and comment coexistence on the same sheet.
- Added the matching promoted Jazzer protocol-request seed plus deterministic Jazzer support
  coverage for drawing-media workflows, so the public example is replay-verified and the nested
  round-trip or invariant layer now asserts the drawing contract directly.

### Changed

- Public docs, quick-reference snippets, README guidance, and the internal XSSF parity and
  inventory records now describe the shipped drawing, image, and embedded-object platform
  explicitly, including authored two-cell anchors, factual read-side anchor variants, and
  drawing-payload extraction boundaries.

### Fixed

- Jazzer sequence labeling, workflow generation, response invariants, and `.xlsx` round-trip
  verification now model the drawing-media command and read surface instead of silently lagging
  behind it.
- The direct-POI parity oracle now stores embedded-object payload bytes behind a defensive
  immutable class, keeping the parity build warning-free under Error Prone's array-component
  checks.

## [0.37.0] - 2026-04-12

### Added

- Added `examples/formula-environment-request.json`,
  a runnable example covering top-level `formulaEnvironment`, template-backed UDF
  registration, targeted formula evaluation, and explicit formula-cache clearing.
- Added the matching promoted Jazzer protocol-request seed for the formula-environment example, so
  the public request now stays replay-verified in the committed regression floor.

### Changed

- Public docs, quick-reference snippets, README guidance, and the internal XSSF parity records now
  describe the completed formula-evaluation contract explicitly: external workbook bindings,
  missing-workbook policy control, template-backed UDF toolpacks, targeted formula evaluation, and
  explicit persisted-cache clearing.

### Fixed

- `CLEAR_FORMULA_CACHES` now clears persisted formula cached results in the workbook while normal
  post-mutation invalidation still resets only the in-process evaluator cache, so explicit
  lifecycle control is honest without changing ordinary mutation parity semantics.
- The direct-POI parity oracle now evaluates external-link and UDF scenarios through read-only
  input streams before applying transient evaluator configuration, so parity measurement no longer
  mutates the corpus workbook it is measuring.
- Jazzer generation, labeling, and `.xlsx` round-trip verification now cover targeted formula
  evaluation and explicit cache clearing, so the fuzz support layer no longer lags behind the
  shipped formula contract.
- Formula-health analysis and request/command dispatch now route the completed formula-lifecycle
  families through explicit type-owned paths, removing stale unreachable fallback branches and
  making the verification surface match the runtime architecture more directly.

## [0.36.0] - 2026-04-11

### Added

- Added `examples/advanced-mutation-request.json`, a
  runnable workbook-core mutation example covering password-bearing protection, formula-defined
  named ranges, advanced table and autofilter mutation, advanced conditional formatting, rich
  comments, advanced page setup, and structured style colors.

### Changed

- Public docs, quick-reference snippets, README guidance, the internal XSSF capability inventory,
  and the internal parity execution spec now describe the completed non-drawing workbook-core
  mutation contract explicitly instead of the earlier partial summaries.
- `examples/advanced-readback-request.json` now
  materializes the richer factual readback surface it advertises, including workbook protection,
  rich comment runs and anchors, advanced page setup, structured style colors and gradients,
  autofilter criteria and sort state, and advanced table metadata.

### Fixed

- `SET_PRINT_LAYOUT` docs now cover the supported advanced page-setup payload instead of only the
  earlier core print-layout subset.
- Public contract docs now describe the real `SET_SHEET_PROTECTION`,
  `SET_WORKBOOK_PROTECTION`, `SET_COMMENT`, `APPLY_STYLE`, `SET_AUTOFILTER`, `SET_TABLE`,
  `SET_CONDITIONAL_FORMATTING`, and `SET_NAMED_RANGE` surfaces, including password-bearing
  protection, rich comments, structured color writes, gradient fills, advanced filter metadata,
  advanced table metadata, six conditional-format rule families, and formula-defined names.
- The internal XSSF parity oracle now detects workbook and revisions password-hash presence across
  both legacy and modern OOXML workbook-protection fields, so parity verification no longer reports
  a false regression on SHA-512-authored workbooks.

## [0.35.0] - 2026-04-11

### Added

- GridGrind now exposes read-parity surface for workbook protection, rich comment runs
  and anchors, advanced print setup, structured theme or indexed or tinted color facts,
  gradient fills, autofilter criteria and sort state, advanced table metadata, and the
  remaining POI-readable XSSF conditional-formatting families modeled by GridGrind.
- Added `examples/advanced-readback-request.json` plus
  the matching promoted Jazzer seed so the richer factual readback contract is both publicly
  demonstrated and replay-verified.

### Changed

- The CLI help now states that column structural edits are blocked by any workbook formulas, not
  just formulas on the sheet being edited.
- The protocol catalog, public docs, README, and the promoted `table_autofilter_request.json`
  Jazzer seed now use the real defaulted `SET_TABLE` contract: omit `showTotalsRow` unless the
  table actually includes a totals row.
- The public docs, quick reference, README, Jazzer docs, and parity records now describe the full
  richer readback contract instead of the earlier narrowed summaries.

### Fixed

- `SET_TABLE` now treats `showTotalsRow` as a genuinely optional request field that defaults to
  `false`, matching the protocol catalog and black-box CLI behavior.
- The protocol catalog builder now rejects any future attempt to mark a primitive record component
  as optional, so catalog optionality cannot drift away from JSON deserialization semantics again.
- Data-validation health analysis now preserves distinct malformed raw-validation states instead of
  collapsing them into broader findings, and row or column structural-edit guards now tolerate
  malformed raw validation records without crashing while computing unsupported-formula checks.
- Advanced readback now degrades malformed raw conditional-formatting family metadata and malformed
  raw data-validation enum metadata into factual unsupported reports instead of crashing before the
  workbook can be inspected.
- Persisted autofilter sort-state and sort-condition ranges are now reported exactly as stored,
  including blank raw ranges, so malformed workbook metadata is surfaced to callers instead of
  being rejected during factual readback.
- The advanced XSSF parity corpus now correctly materializes workbook autofilters, theme or tinted
  font colors, and gradient fills, so the parity oracle measures the intended read surface instead
  of an underspecified fixture subset.

## [0.34.0] - 2026-04-11

### Added

- Added an executable Apache POI `5.5.1` XSSF parity gate. GridGrind now ships a canonical parity
  ledger, a golden `.xlsx` corpus, a direct-POI oracle harness, GridGrind-side comparator probes,
  and a root `./gradlew parity` task for measuring current `.xlsx` parity status end to end.

### Changed

- Root and protocol `check` verification now includes the XSSF parity source set plus parity PMD
  coverage, so `.xlsx` parity drift fails local verification instead of living only in ad hoc
  investigation.
- Jazzer operator docs now distinguish the supported `jazzer/bin/*` surface from raw
  `./gradlew --project-dir jazzer ...` debugging more explicitly, and the seed-inventory docs now
  point at one authoritative exhaustive committed-input list instead of drifting count copies.
- Jazzer operator docs now state a single supported active-fuzz method only: `jazzer/bin/*`.
  Raw Gradle is documented only for deterministic nested-build verification, not as an endorsed
  alternative fuzz entrypoint.

### Fixed

- Style snapshot extraction no longer preserves border colors when the effective border style is
  `NONE`, avoiding impossible border-state reports for advanced POI-authored `.xlsx` workbooks.
- The parity corpus now materializes real agile-encrypted OOXML workbooks, and the direct-POI
  parity oracle opens them through POI's decryptor flow instead of treating them as plain `.xlsx`
  files.
- Nested Jazzer active fuzzing now preloads a project-owned premain agent that publishes startup
  instrumentation to Byte Buddy before Jazzer's JUnit extension runs, so Java 26 live fuzzing no
  longer wedges in the external attach path before the harness starts executing.
- Active Jazzer fuzzing now hard-fails on GitHub Actions, so GitHub remains a deterministic-only
  verification surface even if an active fuzz task is wired there by mistake.
- Promoted Jazzer metadata no longer carries stray committed `.txt.tmp` artifacts, and
  `PromotionMetadataTest` now rejects non-`.json`/`.txt` files plus orphan replay-text artifacts
  so temporary refresh leftovers cannot silently re-enter version control.
- Root `./check.sh` stall diagnostics now bound heavyweight per-process captures to a small sample,
  so a badly wedged stage cannot fan out `lsof` or `jcmd` collection across an unbounded
  descendant process tree.
- Supported `jazzer/bin/*` active-fuzz runs now force `--no-daemon` and tear down the launched
  Gradle client tree on interrupt or timeout, so canceling a local fuzz session no longer drops the
  run lock while leaving a live harness JVM and wrapper client chewing CPU in the background.

## [0.33.0] - 2026-04-10

### Changed

- Root build policy is now fully convention-plugin owned. The repository root build script is
  reduced to a thin `gridgrind.root-conventions` application, shared Java module policy lives in
  `GridGrindJavaConventionsPlugin`, and the nested Jazzer build now inherits the same shared
  Spotless and PMD enforcement instead of depending on an old root `subprojects {}` block.
- Nested Jazzer verification now uses dedicated local-only static-analysis profiles: a Jazzer PMD
  ruleset for support and operator code, a fuzz-harness PMD ruleset for `@FuzzTest` entrypoints,
  and a dedicated JaCoCo verification scope for deterministic support-contract classes rather than
  the root product modules' blanket 100% bundle gate.

### Fixed

- Shared build logic no longer recompiles nondeterministically after local edits. The obsolete
  root `buildSrc` directory is gone, build-logic output directories are cleaned without deleting
  the compiler's classpath root, and Kotlin incremental compilation is disabled for
  `gradle/build-logic` so composite-build recompiles no longer lose sibling helper types.
- Root aggregated coverage now resolves reproducibly under configuration cache. The
  `gridgrind.root-conventions` plugin now declares the root repository needed by the
  `jacocoAggregatedReport` task instead of relying on repository state that used to live in the
  old root build script.
- Nested Jazzer support verification is once again fully live under `./gradlew --project-dir
  jazzer check`: the interrupted `XlsxRoundTripVerifierTest` refactor is completed, Jazzer-only
  PMD findings are enforced through the right profile, and regression replay plus support-test
  pulses remain green under the shared conventions.
- Active fuzz scripts once again execute through Jazzer's real command-line JUnit launcher instead
  of a partial in-repo reimplementation. `JazzerHarnessRunner` now requires exactly one
  `@FuzzTest` per harness class, delegates to Jazzer's official `JUnitRunner`, and honors
  `jazzer.max_duration` and `jazzer.max_executions` during local live runs.
- Excel XMLBeans `sqref` handling now goes through one shared normalizer instead of ad hoc raw
  stream usage. Conditional-formatting and data-validation paths both normalize the same way, and
  engine compilation is free of the lingering unchecked-operation notes in the conditional-
  formatting controller and its tests.

## [0.32.2] - 2026-04-10

### Added

- Added [docs/DEVELOPER_GRADLE.md](./docs/DEVELOPER_GRADLE.md), a developer-facing map of the
  Gradle system that explains the shared included build logic, the nested Jazzer composite build,
  the single version-catalog authority, and the periodic review questions contributors should use
  when revisiting the build architecture.

### Changed

- Root and nested Jazzer Gradle builds now share one included build-logic project under
  `gradle/build-logic`, and the nested Jazzer build now imports the root version catalog instead
  of hardcoding overlapping JUnit, Jackson, Apache POI, Log4j, and Jazzer coordinates locally.
- Jazzer harness and run-target metadata now comes from one committed topology file,
  `jazzer/src/main/resources/dev/erst/gridgrind/jazzer/support/jazzer-topology.json`, which is
  consumed by both the runtime support layer and the nested build's task registration.
- `jazzer/build.gradle.kts` is now a thin plugin application rather than a 683-line mixed build
  script, while the nested build still preserves the same public task names and `jazzer/bin/*`
  operator surface.

### Fixed

- `./check.sh` no longer launches stage logging through a racy temporary FIFO that could be
  unlinked before `tee` opened it. Local release verification therefore no longer intermittently
  fails during later stages with `No such file or directory` even when the underlying stage
  command itself succeeds.
- Deleted build-logic helpers can no longer linger as stale hidden `buildSrc` classes in local
  Gradle state. GridGrind now compiles its shared included build logic from clean class output
  directories on each rebuild, and the obsolete root/nested `buildSrc` builds are gone.
- Jazzer support-test pulses and root-project Gradle test pulses now share one scheduled pulse
  foundation, so heartbeat scheduling, thread naming, and whitespace normalization no longer drift
  independently between the two build layers.

## [0.32.1] - 2026-04-09

### Changed

- `./check.sh` now emits explicit per-stage elapsed times plus a total elapsed time in its final
  summary, so long local verification runs show where wall-clock time went without requiring
  external timing wrappers.
- `./check.sh` now derives its Jazzer regression-target progress total from the live regression
  pulse plan instead of a hardcoded harness count, so adding or removing a replay harness cannot
  silently desynchronize the Stage 2 progress summary from the actual nested build.

### Fixed

- Local Docker smoke verification now includes the legal files copied by the production image.
  `.dockerignore` no longer strips those files from the build context, so local image builds match
  the Dockerfile contract instead of failing at copy time.
- Local Docker smoke now asserts GridGrind's help and version output semantically instead of only
  checking that the container process exits successfully, so release-surface contract drift is
  caught before publication.
- Root-project Gradle test pulses now include scheduled in-flight heartbeats for long-running
  tests instead of only reporting progress after completed tests. Local `./check.sh` quality-gate
  monitoring therefore no longer misclassifies healthy long tests as stalled just because a single
  test method runs quietly for longer than the stall threshold.
- Jazzer regression replay now validates promoted-metadata target keys and referenced artifacts
  before replaying them, and mismatch diagnostics now include richer expectation details plus any
  unexpected-failure stack trace. Corrupted or partially moved promoted metadata therefore fails
  fast as a regression-infrastructure defect instead of surfacing as a less actionable replay
  mismatch later.

## [0.32.0] - 2026-04-09

### Changed

- `--help`, the protocol catalog, and the docs now make two important workbook rules explicit:
  Excel sheet names reject reserved characters, and relative `FILE` hyperlinks are checked
  relative to the saved workbook directory during workbook-health analysis.

### Fixed

- Column grouping no longer routes `GROUP_COLUMNS(..., collapsed=false)` through Apache POI's
  collapsed-group expansion path. Overlapping expanded and previously collapsed column groups now
  stay deterministic instead of surfacing the XMLBeans `IndexOutOfBoundsException` that POI can
  trigger while rewriting split column definitions.
- Column outline edits now discard ghost column metadata and canonicalize ambiguous Excel column
  definitions before layout reads and persistence. Repeated no-op ungroup operations therefore no
  longer poison later collapsed groups, and overlapping outline edits keep the same visible state
  across save/reopen cycles instead of drifting when Apache POI leaves stale column definitions in
  memory.
- Sheet-name validation is now consistent across request parsing and engine execution. Invalid
  Excel sheet-name characters and leading or trailing apostrophes are rejected up front with
  structured `READ_REQUEST` failures instead of leaking raw Apache POI sheet-creation errors
  later in `APPLY_OPERATION`.
- Missing relative `FILE` hyperlink findings now explain the workbook directory they were
  resolved against, so the health-check output makes the relative-path anchor obvious instead of
  forcing callers to reverse-engineer it from the final resolved path alone.
- Jazzer now replays the previously crashing engine-command-sequence artifact
  `overlapping_collapsed_then_expanded_group_columns_expected_invalid` as a committed
  expected-invalid regression input, locking the no-crash behavior into the verification suite.

## [0.31.0] - 2026-04-08

### Added

- Jazzer promoted seed `partial_collapsed_column_ungroup_roundtrip_success` now locks in the
  `.xlsx` round-trip case where a collapsed grouped column band is partially ungrouped after save.
- Jazzer support-test progress pulses now include committed class-start and in-flight
  `test-progress` heartbeats, so long-running nested support tests remain observable to the outer
  gate instead of going dark between test completions.
- Added [docs/DEVELOPER_JAVA.md](./docs/DEVELOPER_JAVA.md), which documents GridGrind's shell-level
  Java 26 setup, the `/usr/bin/java` macOS pitfall, the official `jdk.java.net/26` install path,
  and why `./gradlew` is the only supported Gradle entrypoint.
- Added a no-save workbook-health example at
  [examples/workbook-health-request.json](./examples/workbook-health-request.json), showing
  `ANALYZE_WORKBOOK_FINDINGS` as the default lint-style workflow and the correct quoted formula
  syntax for sheet names with spaces.
- Successful protocol responses can now include a `warnings` array. The first shipped warning flags
  formulas that reference same-request sheet names with spaces without single quotes.

### Changed

- Upgraded PMD from 7.22.0 to 7.23.0 and Error Prone from 2.48.0 to 2.49.0 across the shared
  Gradle analyzer toolchain, so all Java subprojects now compile and lint against the newer
  static-analysis baselines.
- Jazzer `.xlsx` round-trip verification now snapshots persisted expectations from the actual
  pre-save workbook state and uses command replay only to bound which cell snapshots are compared
  after reopen. This removes duplicated sheet-layout modeling from the verifier path.
- Local `./check.sh` Stage 1 verification now feeds its stall monitor with semantic
  `[GRADLE-TEST-PULSE]` progress from Gradle `Test` tasks, so long-running root quality-gate runs
  are tracked by actual test execution instead of a stale `> Task :engine:test` banner.
- GridGrind's developer and release docs now treat shell-level Java 26 as a first-class runtime
  requirement and explicitly distinguish that from the temporary JVM 25 bytecode target used only
  by Gradle `buildSrc` logic while Kotlin lacks direct JVM 26 output.
- Gradle `buildSrc` logic now compiles with the Java 26 toolchain while still emitting JVM 25
  bytecode, so local builds no longer require a separate Java 25 installation just to satisfy the
  temporary Kotlin build-logic ceiling.
- CLI help, the README, and the public protocol docs now surface GridGrind's coordinate split much
  more explicitly: `address` and `range` stay in A1 notation, while `*RowIndex` and
  `*ColumnIndex` fields are zero-based and rendered back with Excel-native equivalents in
  validation messages.
- `ANALYZE_WORKBOOK_FINDINGS` is now documented consistently as GridGrind's primary
  workbook-health check, including the no-save `persistence.type=NONE` workflow.

### Fixed

- Jazzer no longer reports a false `.xlsx` round-trip failure for partially ungrouped collapsed
  column groups. GridGrind now accepts Excel's persisted boundary-column collapsed-marker
  semantics for that case, and artifact-backed replay tests load exact committed fixture bytes
  instead of hand-copied inline Base64.
- `./check.sh` no longer false-stalls healthy Stage 1 root verification during quiet but active
  Gradle test execution, and its pulse output now shows live test-progress facts instead of only
  task-start lines.
- `./check.sh` no longer false-stalls healthy Jazzer Stage 2 support-test runs because the
  `JazzerSupportTestPulseListener` is once again source-owned and emits progress during long
  individual support tests instead of relying on stale compiled build state and completion-only
  pulses.
- `./check.sh` now fails fast when the active shell resolves `java` or `javac` to the macOS
  launcher stubs or to anything other than Java 26, so local verification cannot silently run on
  the wrong ambient runtime.
- `EXECUTE_REQUEST` failure context now preserves the parsed request's `sourceType` and
  `persistenceType` instead of dropping them to `null`.
- Row and column bounds errors now report Excel-native equivalents inline, such as
  `firstRowIndex 5 (Excel row 6)` and `firstColumnIndex 5 (Excel column F)`, across structural
  edits, print-title bands, and readback validation records.

## [0.30.0] - 2026-04-07

### Added

- Style-system expansion: `APPLY_STYLE` and cell-style reads now expose a nested style contract
  with `numberFormat`, `alignment`, `font`, `fill`, `border`, and `protection` groups. The
  shipped style surface now includes text rotation, indentation, cell `locked` and
  `hiddenFormula` flags, per-side border colors, and patterned fills with foreground or
  background colors.
- Rich-text cell authoring: `SET_CELL`, `SET_RANGE`, and `APPEND_ROW` now accept typed
  `RICH_TEXT` values with ordered runs and optional per-run font overrides.
- String cell reads in `GET_CELLS`, `GET_WINDOW`, and `GET_SHEET_SCHEMA` now surface optional
  structured `richText` runs alongside `stringValue`, so authored rich text round-trips through
  the existing cell-introspection surface instead of a separate read family.

### Changed

- Public request examples, README snippets, and committed Jazzer protocol-request seeds now use
  the nested `APPLY_STYLE` JSON shape instead of the old flat style fields, so docs, examples,
  and deterministic replay all reflect the live public contract.
- Jazzer style generation, style-kind coverage telemetry, protocol response invariants, and
  `.xlsx` reopen verification now operate on the full nested style model rather than a flat
  subset.
- Jazzer typed-value generation, protocol invariants, and `.xlsx` reopen verification now assert
  rich-text persistence explicitly, including run ordering, non-empty run text, concatenation
  back to `stringValue`, and effective per-run font facts on read-back.
- Border color patches now require an effective visible border style on the same side, either set
  directly or inherited from `border.all`, instead of tolerating color-only states that Excel does
  not model as a visible border.

### Fixed

- `APPLY_STYLE` border patches that clear a side or `border.all` back to `NONE` no longer crash on
  workbooks whose underlying border XML already has no stored color entry. Border-color clearing
  now avoids POI's unsafe unset path when there is no color to remove, and the case is locked into
  deterministic `.xlsx` round-trip regression coverage.

## [0.29.0] - 2026-04-07

### Added

- Column structural editing:
  `INSERT_ROWS`, `DELETE_ROWS`, `SHIFT_ROWS`, `INSERT_COLUMNS`, `DELETE_COLUMNS`,
  `SHIFT_COLUMNS`, `SET_ROW_VISIBILITY`, `SET_COLUMN_VISIBILITY`, `GROUP_ROWS`,
  `UNGROUP_ROWS`, `GROUP_COLUMNS`, and `UNGROUP_COLUMNS`.
- `GET_SHEET_LAYOUT` now reports row and column `hidden`, `outlineLevel`, and `collapsed`
  facts alongside explicit size so grouped or hidden structure can be read back cleanly.
- New example request: `examples/row-column-structure-request.json`.
- Round-trip and Jazzer verification now assert persisted row or column layout state for column structural
  operations, including hidden-state shifts, grouping, and save-reopen fidelity.
- Jazzer promoted seed `row_column_structure_request` now covers the public structural-edit
  request surface, including row or column insertion, deletion, shifting, visibility, grouping,
  and `GET_SHEET_LAYOUT` reads.

### Changed

- `./check.sh` now emits `[CHECK-PULSE]` stage progress lines during long-running verification
  and captures `[CHECK-DIAG]` stall snapshots with process-tree, `lsof`, thread-dump, and log-tail
  artifacts when a stage stops making semantic progress. Diagnosed stalled stages are now
  terminated with a stable failure instead of waiting indefinitely.
- Nested Jazzer verification now emits stable `[JAZZER-PULSE]` progress lines for deterministic
  support tests, per-harness regression targets, per-input committed-seed replay, and standalone
  harness replay, so long-running Stage 2 verification exposes real progress instead of appearing
  silent.
- Promoted-input semantic replay now runs only in the dedicated Jazzer regression runner tasks
  instead of inside `PromotionMetadataTest`, keeping the support-test JVM structural-only while
  preserving committed replay-contract verification in the isolated tool runtime.
- Deterministic Jazzer replay now uses a project-owned pure-Java scalar fuzz-data cursor for the
  subset of `FuzzedDataProvider` behavior GridGrind actually consumes, so `jazzer/bin/replay`
  and committed-seed regression no longer depend on Jazzer's native replay provider loading in a
  fresh JVM.
- Column structural edits now re-normalize explicit column metadata after insert, delete, and
  shift operations so hidden state, outline state, and other explicit column-definition facts move
  with the authored columns instead of being left behind in stale XML.
- GridGrind now documents and surfaces two product-owned structural-edit limits:
  `LIM-016` rejects edits that would move or truncate tables, sheet-owned autofilters, or data
  validations; `LIM-017` rejects column structural edits when formulas or formula-defined names
  are present because Apache POI leaves some column references stale.

### Fixed

- Row and column destructive structural edits now reject range-backed named ranges before Apache
  POI can rewrite them into broken `#REF!` formulas. `DELETE_ROWS`, `SHIFT_ROWS`,
  `DELETE_COLUMNS`, and `SHIFT_COLUMNS` now preserve workbook integrity by allowing only safe
  named-range cases such as full-band moves and completely untouched ranges.
- Column structural edits now preserve Apache POI's required `<cols>` worksheet container even
  when a sheet has no explicit column metadata, so follow-up operations like `GROUP_COLUMNS`,
  `UNGROUP_COLUMNS`, and column visibility changes no longer crash after insert, delete, or shift
  edits rebuild column definitions.
- Collapsed row groups at the sparse tail of a sheet now persist their control-row marker across
  save and reopen, so `GROUP_ROWS(..., collapsed=true)` round-trips the exposed `collapsed` fact
  in `GET_SHEET_LAYOUT` instead of losing it on reopen.
- Sparse rows touched by `UNGROUP_ROWS` now normalize Apache POI's internal outline-level
  sentinel `-1` back to public outline level `0`, so row layout reads and round-trip assertions
  stay stable after save and reopen.
- Promoted Jazzer metadata and replay text now store project-relative paths instead of
  hard-coded workspace paths. `jazzer/bin/promote`, promoted-metadata refresh, and orphaned-seed
  detection now resolve those paths against the project directory, so replay validation works
  correctly from alternate worktrees.

## [0.28.0] - 2026-04-06

### Changed

- `GridGrindJson.cleanJacksonMessage` now strips Jackson configuration-advice suffixes of the
  form `(set X.Y to 'Z' to allow)` in addition to the existing source-location, subtype, and
  POJO-property noise patterns. This closes a class of raw-message leaks: any null-into-primitive
  coercion on any field type now has its configuration advice stripped before the message reaches
  `productOwnedJacksonMessage`, regardless of whether a dedicated dispatch arm exists for that
  exact exception shape.

### Fixed

- `GridGrindJson` now surfaces a clean `Missing required field '<name>'` error when a required
  primitive boolean field (e.g. `stopIfTrue`) is supplied as JSON `null`. Previously the raw
  Jackson internal configuration-advice message
  (`"set DeserializationConfig.DeserializationFeature.FAIL_ON_NULL_FOR_PRIMITIVES to 'false'
  to allow"`) leaked directly into the `InvalidRequestShapeException` message seen by agents.
- Protocol catalog `DifferentialStyleInput` descriptor now lists all nine optional fields
  (`numberFormat`, `bold`, `italic`, `fontHeight`, `fontColor`, `underline`, `strikeout`,
  `fillColor`, `border`). Previously the field list was empty, so `--describe-type` showed no
  fields at all for this type.
- Protocol catalog descriptions for `FORMULA` cell input, `FORMULA_LIST` and `CUSTOM_FORMULA`
  data-validation rule types, and `FORMULA_RULE` conditional-formatting rule type no longer
  instruct agents to omit the leading `=` sign. GridGrind accepts and strips it automatically;
  the prior wording contradicted the actual behavior observed when submitting `=SUM(...)`.

### Added

- `PromotionMetadataTest.everyInputFileHasPromotionMetadata` asserts that every file committed
  to any harness input directory has a corresponding promoted-metadata entry. Previously only
  the inverse direction was enforced: metadata entries were replayed but no test checked whether
  every input *file* had metadata. A seed hand-dropped into the input directory without running
  `jazzer/bin/promote` would compile, replay in regression mode, but leave no stable contract
  that `PromotionMetadataTest` could assert against. The new test closes that gap for all four
  harnesses simultaneously.
- `JazzerReportSupport.orphanedInputs` and `promotedInputPaths` — new public methods that
  identify input files with no promoted-metadata entry by building a set of all
  `promotedInputPath` values from the metadata tree and diffing against the input directory.
  Used by the new test and by `list-corpus`.
- `list-corpus` (`jazzer/bin/list-corpus`) now surfaces a `WARNING: Seeds Without Promotion
  Metadata` section for any harness that has orphaned input files, listing each file and
  reminding the operator to run `jazzer/bin/promote`. Previously the gap was invisible from
  the operator tooling.
- `GridGrindJsonTest` integration test `wrapsNullPrimitiveBooleanFieldAsInvalidRequestShapeWithFieldName`
  verifies that a full `readRequest` round-trip with `"stopIfTrue": null` produces
  `InvalidRequestShapeException` with message `"Missing required field 'stopIfTrue'"` and no
  Jackson internals in the message.
- Four `GridGrindJsonTest` unit tests for the `mismatchedInputMessage` null-into-primitive paths
  (named property, no path, array-index path) and `cleanJacksonMessage` configuration-advice
  stripping.
- `GridGrindJsonTest` unit test `mismatchedInputMessageWithNullOriginalMessageReturnsFallback`
  covering the `original == null` branch of `mismatchedInputMessage`.
- Jazzer promoted seed `invalid_request_shape_null_primitive_boolean` covering the
  `SET_CONDITIONAL_FORMATTING` request with `"stopIfTrue": null`, captured as
  `EXPECTED_INVALID` / `INVALID_REQUEST_SHAPE`.
- Jazzer promoted seed `invalid_request_shape_null_primitive_int` covering an `APPLY_STYLE`
  request with `"fontHeight": {"type": "TWIPS", "twips": null}`, captured as `EXPECTED_INVALID`
  / `INVALID_REQUEST_SHAPE` with message `"Missing required field 'twips'"`. Proves that the
  `cleanJacksonMessage` configuration-advice stripping generalizes to non-boolean primitives.
- Fourteen previously unpromoted protocol-request seeds retroactively promoted with full
  expectation metadata and replay text: `clear_on_empty_cells`, `delete_last_sheet`,
  `duplicate_request_id`, `formula_equals_prefix`, `get_cells_invalid_address`,
  `get_cells_out_of_bounds_address`, `get_window_overflow`, `introspection_analysis_request`,
  `invalid_email_no_at_sign`, `schema_empty_sheet`, `schema_formula_cells`,
  `sheet_name_too_long`, `unknown_field_rejection`, `window_size_limit_exceeded`. All were
  curated seeds in the corpus without regression contracts; `PromotionMetadataTest` now covers
  them.

## [0.27.0] - 2026-04-05

### Changed

- Eleven shadow enums deleted from `protocol.dto` (`HorizontalAlignment`, `VerticalAlignment`,
  `BorderStyle`, `SheetVisibility`, `PrintOrientation`, `PaneRegion`, `ComparisonOperator`,
  `DataValidationErrorStyle`, `ConditionalFormattingIconSet`, `ConditionalFormattingThresholdType`,
  `ConditionalFormattingUnsupportedFeature`). All protocol DTOs, reports, and operation types now
  reference the canonical `Excel*` engine enums directly via the existing module dependency.
- `DefaultGridGrindRequestExecutor` (2 571 lines) decomposed into three package-private converter
  classes: `WorkbookCommandConverter` (protocol operations to engine commands),
  `WorkbookReadCommandConverter` (protocol read operations to engine read commands), and
  `WorkbookReadResultConverter` (engine read results to protocol reports). The executor now
  delegates to these converters and contains no conversion logic. Each converter exposes
  package-private static methods for direct unit testing without reflection.

### Added

- New Jazzer protocol-request regression seed `conditional_formatting_request` covering
  `SET_CONDITIONAL_FORMATTING` (formula rule, cell-value rule with `LESS_THAN`, cell-value rule
  with `BETWEEN` including a differential border), `CLEAR_CONDITIONAL_FORMATTING` with a selected
  range, `GET_CONDITIONAL_FORMATTING`, and `ANALYZE_CONDITIONAL_FORMATTING_HEALTH`. No seed
  exercising conditional-formatting operations existed previously.

### Fixed

- Promoted protocol-request seed `conditional_formatting_request` uses `conditionalFormatting`
  for `SET_CONDITIONAL_FORMATTING` payloads, matching `WorkbookOperation.SetConditionalFormatting`
  JSON and Jazzer promotion-metadata replay expectations.

## [0.26.0] - 2026-04-04

### Added

- New Jazzer protocol-request regression seed `pane_and_print_reset_request` covering the
  previously untested pane and print-layout variants: `PaneInput.NONE` (reset), `PaneInput.SPLIT`
  (x/y offsets with active-pane region), `PrintAreaInput.NONE`, `PrintScalingInput.AUTOMATIC`,
  and `CLEAR_PRINT_LAYOUT`. The existing `structural_layout_request` seed only covered `FROZEN`
  panes with `FIT` scaling.
- Protocol catalog now publishes `paneTypes`, `printAreaTypes`, `printScalingTypes`,
  `printTitleRowsTypes`, and `printTitleColumnsTypes` as nested type groups, and
  `headerFooterTextInputType` and `printLayoutInputType` as plain type groups. Agents using
  `--print-protocol-catalog` or `--describe-operation SET_SHEET_PANE` / `SET_PRINT_LAYOUT`
  previously received dangling `NESTED_TYPE_GROUP` or `PLAIN_TYPE_GROUP` references with no
  matching catalog entry; those entries are now present and complete.
- Bidirectional validation between the field-shape group maps in `CatalogFieldMetadataSupport`
  and the descriptor lists in `GridGrindProtocolCatalog`: a type registered in the field-shape
  map with no corresponding catalog descriptor now raises `IllegalStateException` at startup
  rather than silently producing an incomplete catalog.

### Changed

- Eight public nested types (`Catalog`, `TypeEntry`, `FieldEntry`, `NestedTypeGroup`,
  `PlainTypeGroup`, `FieldShape`, `FieldRequirement`, `ScalarType`) extracted from
  `GridGrindProtocolCatalog` to individual top-level files in the `catalog` package.
  Wire format, catalog content, and public API are unchanged.
- Protocol-catalog construction now uses a small internal gatherer seam for ordered uniqueness
  and reflected field-metadata expansion instead of hand-rolled duplicate and ordering logic
  inside `GridGrindProtocolCatalog`.
- Built-in discovery output remains deterministic and contract-identical, while request-template
  generation intentionally stays a plain constant because it does not warrant gatherer-based
  abstraction.
- Release and container workflows now support tag-targeted `workflow_dispatch` reruns, and the
  release procedure now codifies protected-`main` CI requirements plus automatic release-branch
  cleanup.
- GHCR publication verification now runs both `docker pull` and `docker run` through the same
  anonymous Docker config, matching the operator release protocol instead of silently falling
  back to ambient credentials.

## [0.25.0] - 2026-04-03

### Changed

- GridGrind now compiles and verifies on the JPMS module path across `engine`, `protocol`, and
  `cli`, enforcing the intended `cli -> protocol -> engine` dependency graph in normal builds.
- The protocol implementation is now split into protocol-owned packages by responsibility:
  `dto`, `operation`, `read`, `catalog`, `exec`, and `json`, instead of the older flat
  package layout.
- `DefaultGridGrindRequestExecutor` is now the sole engine-aware class in protocol main source.
  Generic problem construction, JSON handling, and catalog generation remain protocol-owned.

## [0.24.0] - 2026-04-03

### Added

- Structural-layout public surface:
  `SET_SHEET_PANE`, `SET_SHEET_ZOOM`, `SET_PRINT_LAYOUT`, `CLEAR_PRINT_LAYOUT`, and
  `GET_PRINT_LAYOUT`.
- New engine seams for sheet-view and print-layout control, including explicit pane-state,
  zoom, print-area, orientation, fit scaling, repeating rows or columns, and plain
  header/footer text support.

### Changed

- `GET_SHEET_LAYOUT` now reports a generalized `pane` family plus effective `zoomPercent`
  instead of the older freeze-only layout contract.
- The machine-readable protocol catalog, public docs, and Jazzer readable seeds now describe the
  generalized pane and print-layout contract rather than the superseded `FREEZE_PANES` shape.

### Fixed

- Jazzer read introspection and `.xlsx` round-trip structural invariants now treat pane state as
  a generalized workbook-view concern instead of a freeze-only special case.

## [0.23.0] - 2026-04-02

### Added

- Sheet-management public surface:
  `COPY_SHEET`, `SET_ACTIVE_SHEET`, `SET_SELECTED_SHEETS`, `SET_SHEET_VISIBILITY`,
  `SET_SHEET_PROTECTION`, and `CLEAR_SHEET_PROTECTION`.
- `GET_WORKBOOK_SUMMARY` now reports the typed `EMPTY` versus `WITH_SHEETS` workbook-summary
  shape, and non-empty workbooks expose `activeSheetName` plus `selectedSheetNames`.
- `GET_SHEET_SUMMARY` now reports `visibility` and typed sheet-protection state alongside the
  existing structural row and column facts.
- New public example `examples/sheet-management-request.json` covering sheet copy, active and
  selected sheet state, visibility, protection, workbook summary, and sheet summary reads.

### Changed

- Sheet-copy execution is now GridGrind-owned instead of delegating to Apache POI's raw
  `cloneSheet()` behavior. Copying preserves supported sheet-local content while rejecting
  unsupported copy cases such as tables and sheet-scoped formula-defined named ranges.
- Jazzer sequence generation now uses a stable byte-selector grammar for workflow, command, and
  read-family dispatch, so expanding the authored surface no longer requires mutating bounded
  selector ranges in place.
- The committed `sheet_management_request` Jazzer protocol-request seed now exercises the shipped
  sheet-state contract instead of the older rename and move only slice.

### Fixed

- `DELETE_SHEET` now shares the same visible-sheet invariant as `SET_SHEET_VISIBILITY`, so
  deleting the last visible sheet returns `INVALID_REQUEST` instead of crashing during workbook
  view-state normalization.
- Workbook and sheet summaries now expose active-sheet, selected-sheet, visibility, and
  protection state through the public response model instead of truncating the new facts at
  the engine boundary.
- Sheet-management copy and workbook-view normalization now have direct round-trip and fuzz-backed
  verification, including empty-sheet copies, protected-sheet copies, and invalid active-tab
  repair paths.
- `CLEAR_SHEET_PROTECTION` is now idempotent on already unprotected sheets instead of delegating
  into an Apache POI removal path that could throw during `.xlsx` round-trip fuzzing.
- Jazzer promotion metadata is now refreshed against the stable selector grammar, so replay
  expectations remain truthful after the sheet-management expansion instead of silently
  describing a previous generator contract.

## [0.22.0] - 2026-04-02

### Added

- Conditional formatting public surface:
  `SET_CONDITIONAL_FORMATTING`, `CLEAR_CONDITIONAL_FORMATTING`,
  `GET_CONDITIONAL_FORMATTING`, and `ANALYZE_CONDITIONAL_FORMATTING_HEALTH`.
- New public example `examples/conditional-formatting-request.json` covering block authoring,
  factual reads, conditional-formatting health, and aggregate workbook findings.

### Changed

- `ANALYZE_WORKBOOK_FINDINGS` now aggregates conditional-formatting findings alongside formula,
  data-validation, autofilter, table, hyperlink, and named-range findings.

### Fixed

- Protocol coverage, executor mapping coverage, and `.xlsx` round-trip Jazzer invariants now
  assert conditional-formatting authoring and persistence instead of leaving the new family
  under-verified.
- Table header mutations made after `SET_TABLE`, including header-range style patches that change
  typed header display text, now synchronize the persisted table-column metadata immediately.
  `ExcelWorkbook.save()` also normalizes every table header before persistence as a backstop, so
  `GET_TABLES`, table health analysis, save/reopen behavior, and `.xlsx` round-trip fuzzing all
  observe one converged table-header state instead of drifting between sheet cells and table XML.
- Jazzer now has promoted `.xlsx` round-trip success seeds covering both the former table-header
  rewrite crash and the later typed-header style-display crash, and the deterministic round-trip
  verifier suite asserts that header rewrites, header clears, and header-range style changes
  survive save and reopen without metadata drift.

## [0.21.0] - 2026-04-01

### Added

- Table and autofilter public surface:
  `SET_AUTOFILTER`, `CLEAR_AUTOFILTER`, `SET_TABLE`, `DELETE_TABLE`, `GET_AUTOFILTERS`,
  `GET_TABLES`, `ANALYZE_AUTOFILTER_HEALTH`, and `ANALYZE_TABLE_HEALTH`.
- New public example `examples/table-autofilter-request.json` covering sheet-level autofilters,
  table authoring, factual reads, and both health-analysis families.
- Jazzer now has promoted seeds for a readable table-plus-autofilter request, a protocol
  workflow dominated by autofilter behavior, and an `.xlsx` round-trip invalid-table case.
- Jazzer promotion metadata now carries a stable replay expectation contract, and the new
  `jazzer/bin/refresh-promoted-metadata` command refreshes committed replay metadata after
  intentional generator or replay-shape changes.

### Changed

- `ANALYZE_WORKBOOK_FINDINGS` now aggregates autofilter and table findings alongside formula,
  data-validation, hyperlink, and named-range findings.

### Fixed

- Table and autofilter logic now share a dedicated sheet-structure support seam instead of
  coupling controllers through package-private helper methods.
- Jazzer promoted-seed verification no longer drifts silently when sequence-generation behavior
  changes; deterministic support tests now replay every promoted metadata entry and assert that
  its stored replay expectation still matches reality.

## [0.20.0] - 2026-04-01

### Added

- Data validation public surface:
  `SET_DATA_VALIDATION`, `CLEAR_DATA_VALIDATIONS`, `GET_DATA_VALIDATIONS`, and
  `ANALYZE_DATA_VALIDATION_HEALTH`.
- New public example `examples/data-validation-request.json` covering validation authoring,
  partial clearing, factual reads, and health analysis.
- Jazzer now has a promoted valid validation workflow seed and a promoted expected-invalid
  validation seed, so the committed regression floor covers both supported and rejected
  validation shapes.

### Changed

- `ANALYZE_WORKBOOK_FINDINGS` now aggregates data-validation findings alongside formula,
  hyperlink, and named-range findings.
- Protocol discovery now publishes `rangeSelectionTypes`, `dataValidationRuleTypes`,
  `dataValidationInputType`, `dataValidationPromptInputType`, and
  `dataValidationErrorAlertInputType`, so black-box consumers can author the full validation
  workflow without shape inference.
- Developer docs, the Jazzer coverage inventory, and the Apache POI parity inventory now treat
  data validation as shipped behavior instead of a planned gap.

### Fixed

- `.xlsx` round-trip verification, executor integration coverage, and Jazzer invariants now assert
  that normalized data-validation state survives save and reopen instead of silently ignoring the
  new command family.
- `.xlsx` round-trip verification no longer duplicates stale command-semantics assumptions for
  style and metadata persistence. The verifier now snapshots expected pre-save workbook state from
  the actual in-memory workbook, so `APPEND_ROW` date-time writes onto styled blank rows replay
  cleanly and stay covered by the committed Jazzer seed floor.
- `GET_DATA_VALIDATIONS` now exposes only observable public states: `SUPPORTED` and
  `UNSUPPORTED`. Invalid workbook validation structures that Apache POI refuses to materialize are
  no longer advertised as a separate `MALFORMED` entry family, so the read contract matches the
  real workbook-loading seam.

## [0.19.0] - 2026-04-01

### Changed

- `--print-protocol-catalog` now publishes field descriptors instead of loose field-name lists.
  Every catalog entry states whether a field is required or optional and exposes the exact scalar,
  list, nested-group, or plain-group shape accepted by that field.
- The `Release` and `Container` workflows now serialize publication per workflow and tag ref,
  verify the external GitHub release and GHCR handoff after publishing, and treat publication as
  a converged public state instead of a one-shot side effect.
- `./check.sh` now syntax-checks the release-surface shell scripts before the Docker smoke stage,
  so publication helpers fail fast locally and in CI when a shell edit breaks the release path.

### Fixed

- CLI and docs discovery guidance now explain that the machine-readable protocol catalog is the
  authoritative black-box contract for field requirements and polymorphic field shapes, so
  operations such as `SET_HYPERLINK`, `SET_RANGE`, and the selection-based reads can be authored
  without inference.
- Duplicate tag-triggered release runs no longer fail spuriously with `Release.tag_name already
  exists`; GitHub Release publication is now idempotent and asset-safe under duplicate dispatch.
- Duplicate or delayed tag-triggered container publication now has a built-in pull-and-run
  verification step, so the workflow confirms the exact version tag and `latest` are both
  publicly runnable before cleanup proceeds.

## [0.18.0] - 2026-03-31

### Added

- `./check.sh` now runs a fourth stage, `scripts/docker-smoke.sh`, which builds the local Docker
  image and verifies `--help`, `--version`, request-file loading, response-file writing, and
  `.xlsx` persistence from a non-default working directory using weird path names.
- CI now includes a separate `Docker smoke` job, and the `Container` workflow runs the same smoke
  script before publishing multi-arch images to GHCR.

### Changed

- `FILE` hyperlinks continue the hard-break path contract: requests use `FILE.path`, plain paths
  and `file:` URIs normalize to plain paths, read surfaces return plain paths, and hyperlink
  health resolves relative file targets against the workbook location.
- `APPEND_ROW` continues to use value-bearing row semantics, so metadata-only rows do not shift
  the append cursor.
- `AUTO_SIZE_COLUMNS` continues to use deterministic content-based sizing so container and local
  runs agree on column widths.
- CLI `--help`, protocol catalog discovery, `README.md`, and the public reference docs now
  explain the file workflow explicitly: stdin vs `--request`, stdout vs `--response`, `SAVE_AS`
  vs `OVERWRITE`, current-working-directory path resolution, and Docker `-w /workdir` usage.

### Fixed

- The Docker image remains workdir-independent: `docker run -w /workdir ... --help` and
  file-backed request/response flows now have automated regression coverage.
- Hyperlink health continues to report missing local file targets instead of silently treating
  them as healthy.
- Invalid request-shape and invalid cell-address failures continue to expose product-owned
  diagnostics instead of parser, POI, or Java implementation detail.
- Typed value writes continue to preserve existing style, hyperlink, and comment state, and
  `DATE` / `DATE_TIME` writes continue to merge their formats onto the existing style instead of
  replacing it.

## [0.17.0] - 2026-03-31

### Fixed

- Response-side hyperlink payloads now reuse the canonical discriminated hyperlink shape used by
  `SET_HYPERLINK`. `GET_HYPERLINKS`, `GET_CELLS`, and `GET_WINDOW` now return `FILE` targets in
  the `path` field instead of leaking the legacy `target` field on read.
- Failure-response `causes` entries now expose stable GridGrind problem codes and product-owned
  messages instead of raw Java exception class names or parser-library internals. This is a
  **wire format breaking change** for any client that was matching on the old `type` or
  `className` fields.
- `INVALID_REQUEST_SHAPE` now returns product-owned messages for missing required fields as well,
  so shape failures no longer leak parser or Java type metadata through that branch either.
- Typed value writes now preserve existing cell style, hyperlink, and comment state instead of
  resetting presentation when `SET_CELL`, `SET_RANGE`, or `APPEND_ROW` overwrites a styled blank
  cell.
- `DATE` and `DATE_TIME` writes now merge their required number formats onto the existing cell
  style instead of replacing fill, border, font, alignment, or wrap state.
- Jazzer and deterministic round-trip coverage now assert style preservation when `APPEND_ROW`
  reuses styled blank rows under value-bearing append semantics.

## [0.16.0] - 2026-03-31

### Fixed

- The Docker image entrypoint now uses an absolute JAR path, so `docker run ... -w /any/path`
  works reliably instead of failing before GridGrind starts with `Unable to access jarfile`.
- `ANALYZE_HYPERLINK_HEALTH` now reports missing local file targets and unresolved relative file
  targets instead of silently treating those cases as healthy.
- `INVALID_REQUEST_SHAPE` messages are now product-owned and concise. Unknown fields, unknown
  type values, and wrong token shapes no longer leak Jackson or Java class names into the public
  response.

### Changed

- `FILE` hyperlink targets are now written with the field name `path` instead of `target`. The
  write contract accepts either plain file paths or `file:` URIs, and all read surfaces return
  normalized plain path strings. This is a **wire format breaking change** for clients that still
  send or expect `target` on `FILE` hyperlinks.
- `APPEND_ROW` now appends after the last value-bearing row. Rows that contain only style,
  comment, or hyperlink metadata no longer shift the append cursor.
- `AUTO_SIZE_COLUMNS` now uses deterministic content-based sizing instead of host font metrics,
  so headless, Docker, and local runs produce the same column widths.
- `./check.sh` now runs nested Jazzer `check` after the root quality gates and before CLI fat-JAR
  packaging, giving the one-command local gate deterministic support-test and committed-seed
  regression coverage as well.

### Added

- New public example `examples/file-hyperlink-health-request.json` showcasing `FILE.path`,
  `file:` URI normalization, hyperlink metadata reads, and hyperlink-health analysis.

## [0.15.0] - 2026-03-31

### Fixed

- Number cells in `GET_CELLS` and `GET_WINDOW` responses now return `declaredType` and
  `effectiveType` as `NUMBER` instead of `NUMERIC`. This is a **wire format breaking change**:
  any client that matched on `"NUMERIC"` must be updated to `"NUMBER"`. The value `NUMERIC` was
  an Apache POI internal enum name leaking into the wire vocabulary; `NUMBER` matches the input
  side and is self-explanatory.
- `SET_CELL` and `SET_RANGE` now accept a leading `=` in `FORMULA` cell input. Previously, a
  formula like `"=SUM(A1:A3)"` was sent to Apache POI with the `=` retained, causing an
  `InvalidFormulaException`. The leading `=` is now stripped automatically before handing the
  expression to the engine.
- `DELETE_SHEET` now returns `INVALID_REQUEST` when the operation would delete the last remaining
  sheet in the workbook. Previously the request was forwarded to Apache POI, which threw an
  unclassified `IllegalStateException`; the error is now proactively detected and surfaced with a
  clear message before any workbook state is modified.
- Validation error messages for `SET_ROW_HEIGHT` and `SET_COLUMN_WIDTH` now include both the
  enforced limit and the supplied value so the caller can identify the violation without
  re-reading the protocol catalog. For example: `"heightPoints must not exceed 1638.35 (Excel
  storage limit: 32767 twips): got 2000.0"` instead of `"must be less than or equal to 1638.35"`.

### Changed

- `--print-protocol-catalog` now accepts an optional `--operation <id>` flag. When supplied, the
  output contains only the single catalog entry matching that operation ID. Unknown IDs produce
  an `INVALID_ARGUMENTS` error. Without `--operation`, the full catalog is returned unchanged.
- `APPLY_STYLE` catalog summary now documents the write vs. read shape asymmetry for borders:
  the `border` write object (with `all`, `top`, `right`, `bottom`, `left` sub-objects) is not
  mirrored in the cell snapshot; read responses use flat top-level fields `topBorderStyle`,
  `rightBorderStyle`, `bottomBorderStyle`, and `leftBorderStyle`.
- `DELETE_SHEET` catalog summary now states that deleting the last sheet in a workbook returns
  `INVALID_REQUEST`.
- `DATE` and `DATE_TIME` cell input summaries in `--print-protocol-catalog` and `--help` now
  correctly state that `GET_CELLS` returns `declaredType=NUMBER` (not `declaredType=NUMERIC`) for
  these cells. The help text `Limits:` section is updated accordingly.

## [0.14.0] - 2026-03-30

### Fixed

- `GET_CELLS` and `GET_WINDOW` now reject cell addresses that exceed the Excel 2007 sheet
  boundary (row > 1,048,575 or column > 16,383). Previously, addresses like `XFE1` (column
  16,384) were accepted by `CellReference` with non-negative indices and returned a blank
  snapshot instead of failing with `INVALID_CELL_ADDRESS`.
- `GET_WINDOW` now rejects window dimensions that would extend the window beyond the Excel
  2007 sheet boundary. Previously, a window starting at a valid address could silently overflow
  if `topLeft.row + rowCount - 1 > 1,048,575` or `topLeft.col + columnCount - 1 > 16,383`.
- Row height read-back from `GET_SHEET_LAYOUT` now returns an exact value for heights stored as
  twips. Previously, `row.getHeightInPoints()` returned a `float` which introduced floating-point
  imprecision (e.g., 1,638.35 points stored as 32,767 twips read back as 1,638.3499755859375).
  Heights are now read as integer twips divided by 20.0, eliminating the imprecision.
- Error messages from Jackson for unknown `type` discriminators no longer include internal
  fully-qualified class names (e.g., `dev.erst.gridgrind.contract.read.WorkbookReadOperation`) or
  Jackson-internal POJO property annotations. The message now contains only the unknown
  discriminator value and the list of known type IDs.
- `MOVE_SHEET` error message for an out-of-range `targetIndex` now clearly states the workbook's
  sheet count and the valid 0-based index range. Previously the message said
  "between 0 and N (inclusive)" without clarifying what N represented.

### Changed

- `GET_SHEET_SCHEMA` now counts formula cells by their evaluated result type (NUMERIC, STRING,
  BOOLEAN, ERROR) in `observedTypes` and `dominantType`, rather than as FORMULA. This makes
  the schema reflect the data a consumer actually reads, not the cell's internal storage type.

## [0.13.0] - 2026-03-30

### Fixed

- `GET_CELLS` no longer silently returns a blank cell snapshot for a malformed address (e.g.
  `BADADDR`, `A0`). Requests containing any address that Apache POI cannot resolve to valid row
  and column indices now fail immediately with `INVALID_CELL_ADDRESS`. Previously, `CellReference`
  returned row index `-1` for unparseable addresses, and the engine treated that as an absent cell
  and returned a `BLANK` snapshot with no error.
- Row height validation now enforces the exact documented boundary of 1,638.35 points
  (32,767 twips). Previously, `Math.round`-based twips conversion accepted values up to
  approximately 1,638.37 because `Math.round(1638.37 × 20) = 32767`. The validation now uses a
  direct floating-point comparison `heightPoints > Short.MAX_VALUE / 20.0` so that any value above
  1,638.35 is rejected regardless of rounding.

### Changed

- All response `type` discriminator values now echo the corresponding request `type` exactly. This
  is a **wire format breaking change** for any client that inspects the `type` field in read results
  or persistence outcomes. The full mapping from old to new discriminator values:

  Read result discriminators:
  - `WORKBOOK_SUMMARY` → `GET_WORKBOOK_SUMMARY`
  - `NAMED_RANGES` → `GET_NAMED_RANGES`
  - `SHEET_SUMMARY` → `GET_SHEET_SUMMARY`
  - `CELLS` → `GET_CELLS`
  - `WINDOW` → `GET_WINDOW`
  - `MERGED_REGIONS` → `GET_MERGED_REGIONS`
  - `HYPERLINKS` → `GET_HYPERLINKS`
  - `COMMENTS` → `GET_COMMENTS`
  - `SHEET_LAYOUT` → `GET_SHEET_LAYOUT`
  - `FORMULA_SURFACE` → `GET_FORMULA_SURFACE`
  - `SHEET_SCHEMA` → `GET_SHEET_SCHEMA`
  - `NAMED_RANGE_SURFACE` → `GET_NAMED_RANGE_SURFACE`
  - `FORMULA_HEALTH` → `ANALYZE_FORMULA_HEALTH`
  - `HYPERLINK_HEALTH` → `ANALYZE_HYPERLINK_HEALTH`
  - `NAMED_RANGE_HEALTH` → `ANALYZE_NAMED_RANGE_HEALTH`
  - `WORKBOOK_FINDINGS` → `ANALYZE_WORKBOOK_FINDINGS`

  Persistence outcome discriminators:
  - `NOT_SAVED` → `NONE`
  - `SAVED_AS` → `SAVE_AS`
  - `OVERWRITTEN` → `OVERWRITE`

  With symmetric naming, the response `type` field directly identifies which read or persistence
  operation produced it, eliminating any need for a client-side translation table.
- `GET_WORKBOOK_SUMMARY` catalog summary now states that the response includes the list of sheet
  names in the workbook.
- `GET_CELLS` catalog summary now documents that invalid cell addresses (malformed or out-of-range)
  produce `INVALID_CELL_ADDRESS`, and that `effectiveType` is always `FORMULA` for formula cells
  regardless of the evaluated result type.
- `GET_WINDOW` catalog summary now notes that `effectiveType` is always `FORMULA` for formula
  cells.
- Persistence outcome summaries in `--print-protocol-catalog` now state that the response `type`
  echoes the request persistence `type` exactly (e.g. `NONE` persistence returns `type: "NONE"`).
- The container cleanup workflow now prunes GHCR package versions through GitHub's Packages API
  via `gh api`, anchored to the five newest tagged releases, instead of using the stale
  `actions/delete-package-versions` wrapper. This removes the Node20 deprecation warning and
  keeps complete multi-arch release groups together even when GitHub emits multiple untagged
  platform and attestation manifests per release.

## [0.12.0] - 2026-03-29

### Fixed

- `SAVE_AS` and `OVERWRITE` persistence now normalize the save path to its absolute canonical
  form before writing. Paths containing `..` segments (e.g. `/workdir/../out.xlsx`) are resolved
  to their canonical equivalents (`/out.xlsx`) before the file is written. `executionPath` in
  the response now always reflects the true path on disk. Previously, `..` segments were
  preserved in `executionPath` and the file was written to the un-normalized location.
- `GET_WINDOW` and `GET_SHEET_SCHEMA` now reject requests where `rowCount * columnCount` exceeds
  250,000 cells with `INVALID_REQUEST` before any workbook work occurs. Previously, large windows
  (e.g., 1000x1000) could crash the process with `OutOfMemoryError` and produce an empty response
  file. The 250,000-cell limit is a GridGrind operational constraint (not an Excel or Apache POI
  limit; Excel supports up to 1,048,576 rows x 16,384 columns) calibrated to prevent heap
  exhaustion during JSON response serialization in bounded-heap container environments.
- `CLEAR_RANGE` is now a no-op on rows and cells that do not physically exist. Previously it
  materialized phantom rows and cells into the sheet, inflating `physicalRowCount` and distorting
  `GET_SHEET_SUMMARY` results.
- `CLEAR_HYPERLINK` and `CLEAR_COMMENT` are now no-ops when the target cell does not physically
  exist, matching the idempotent behavior of `CLEAR_RANGE`. Previously they returned
  `CELL_NOT_FOUND`.
- `GET_SHEET_SCHEMA` now returns `dataRowCount = 0` when every cell in the inferred header row is
  blank. Previously it returned `rowCount - 1` even for empty sheets with no header data.
- `--request` and `--response` that resolve to the same path are now rejected at argument parse
  time with `INVALID_ARGUMENTS`. Previously the response write silently overwrote the request.

### Changed

- `Execution:` section of `--help` now reads "saves the workbook (unless persistence is NONE)"
  instead of the ambiguous word "persistence", which could be confused with the `persistence`
  JSON field.
- `--help` now includes a `Limits:` section listing all hard constraints upfront: `.xlsx`-only
  format, 31-character sheet names, 250,000-cell window cap, 255-unit column width ceiling,
  1,638-point row height ceiling, and the `DATE`/`DATE_TIME` write-only note. Agents and users
  can now read every hard constraint before constructing any request.
- `--help` now explicitly states that a NEW workbook starts with zero sheets and that
  `ENSURE_SHEET` must be used to create the first sheet.
- `ENSURE_SHEET` and `RENAME_SHEET` summaries in `--print-protocol-catalog` now state the
  31-character sheet name limit.
- `SET_COLUMN_WIDTH` summary in `--print-protocol-catalog` now states `widthCharacters` must
  be > 0 and ≤ 255.
- `SET_ROW_HEIGHT` summary in `--print-protocol-catalog` now states `heightPoints` must be > 0
  and ≤ 1,638.35 (32,767 twips).
- `DATE` and `DATE_TIME` cell input summaries in `--print-protocol-catalog` now note that these
  are write-only type hints stored as Excel serial numbers; `GET_CELLS` returns
  `declaredType=NUMERIC` with a formatted `displayValue`.
- `CLEAR_HYPERLINK` and `CLEAR_COMMENT` summaries in `--print-protocol-catalog` updated to
  reflect the no-op behavior on non-existent cells.
- `GET_WINDOW` and `GET_SHEET_SCHEMA` summaries in `--print-protocol-catalog` now state the
  250,000-cell limit.
- `GET_SHEET_SCHEMA` summary notes that `dataRowCount` is 0 when the header row is entirely blank.
- `NEW` source type summary now notes that a new workbook starts with zero sheets.
- `SAVE_AS` persistence type summary in `--print-protocol-catalog` now documents `requestedPath`
  (the literal path from the request) vs `executionPath` (the absolute normalized path where the
  file was written), and states that missing parent directories are created automatically.
- `GET_SHEET_SUMMARY` summary in `--print-protocol-catalog` now states the semantics of
  `physicalRowCount` (sparse materialized row count), `lastRowIndex` (0-based, -1 when empty),
  and `lastColumnIndex` (0-based, -1 when empty).
- `GET_CELLS` summary now documents the cell snapshot response shape: `address`, `declaredType`,
  `effectiveType`, `displayValue`, `style`, `metadata`, and type-specific value fields
  (`stringValue`, `numberValue`, `booleanValue`, `errorValue`, `formula`, `evaluation`).
- `GET_CELLS` summary now explicitly states that `style.fontHeight` in read responses is a plain
  object with both `twips` and `points` fields, not the discriminated `FontHeightInput` write
  format. `GET_WINDOW` summary cross-references the `GET_CELLS` cell snapshot shape.
- `fontHeightTypes` entries in `--print-protocol-catalog` now document the write format and the
  read-back shape asymmetry so agents can round-trip font height without format confusion.

### Added

- Protocol catalog `TypeEntry` now includes a `fieldEnumValues` map enumerating valid string
  values for fields that accept a finite enumerated set. `CellStyleInput.horizontalAlignment`
  lists all `ExcelHorizontalAlignment` values, `CellStyleInput.verticalAlignment` lists all
  `ExcelVerticalAlignment` values, and `CellBorderSideInput.style` lists all `ExcelBorderStyle`
  values. Agents can now discover valid alignment and border-style values from the catalog
  without trial-and-error.
- `docs/OPERATIONS.md` now includes a "Cell snapshot shape" subsection under `GET_CELLS`
  documenting all common and type-specific fields, the `fontHeight` read-vs-write asymmetry,
  and a field-level table for `GET_SHEET_SUMMARY` response semantics (`physicalRowCount`,
  `lastRowIndex`, `lastColumnIndex`). The `SAVE_AS` section now documents `requestedPath` vs
  `executionPath`.
- `docs/LIMITATIONS.md`: new reference document structured as a numbered registry (`LIM-001`
  through `LIM-015`). Each entry carries a stable ID, the enforced limit value, the error code
  and message raised on violation, the applicable operations, a code reference, and a UX
  reference. Covers the GridGrind operational window limit (250,000 cells), all Excel/Apache POI
  structural limits (rows, columns, text length, cell styles, hyperlinks, formula length,
  nested functions, function arguments), and protocol-level limits (sheet name length, column
  width, row height). Links to Apache POI `SpreadsheetVersion` apidocs and the Microsoft Excel
  specifications page.
- All limit enforcement sites in source code now carry a trailing `// LIM-NNN` comment
  cross-referencing the corresponding registry entry in `docs/LIMITATIONS.md`: `LIM-001` on
  `MAX_WINDOW_CELLS` in both protocol and engine, `LIM-002` on the `.xlsx` path check,
  `LIM-003` on the sheet name length check, `LIM-004` on the column width check, `LIM-005`
  on the row height check, `LIM-006` on the duplicate read request ID check, and `LIM-007` on
  the `GET_CELLS` address list validation.

## [0.11.0] - 2026-03-28

### Added

- `--help` output now includes a one-line product description sourced from `gradle.properties`
  as the single canonical definition.
- Protocol catalog now enumerates five plain (non-polymorphic) request types:
  `CommentInput`, `NamedRangeTarget`, `CellStyleInput`, `CellBorderInput`, and
  `CellBorderSideInput`, with their required and optional fields listed.
- Sheet name length validation: `ENSURE_SHEET`, `RENAME_SHEET`, `SET_NAMED_RANGE`, and every
  other operation that accepts a `sheetName` now rejects names longer than 31 characters with
  `INVALID_REQUEST` before any workbook state is touched.
- Duplicate `requestId` detection in `GridGrindRequest`: construction fails at protocol
  deserialization time when two or more reads share the same `requestId`.
- Email hyperlink address validation: `HyperlinkTarget.Email` now rejects addresses that lack
  `@`, have an empty local part, or have an empty domain with `INVALID_REQUEST`.

### Changed

- Unknown JSON fields in requests are now rejected with `INVALID_REQUEST_SHAPE` instead of
  being silently ignored (`FAIL_ON_UNKNOWN_PROPERTIES` enabled on the protocol mapper).
- Mutation operations (`SET_CELL`, `SET_RANGE`, `APPLY_STYLE`, `SET_HYPERLINK`, `SET_COMMENT`,
  `APPEND_ROW`, `AUTO_SIZE_COLUMNS`) now require the target sheet to exist; they no longer
  auto-create it. Use `ENSURE_SHEET` before the first write to a sheet.
- `GET_CELLS` now returns a blank-typed cell snapshot for addresses that have never been written
  rather than returning `CELL_NOT_FOUND`. Empty cells are valid cells.
- `FORMULA` cell input type description in `--print-protocol-catalog` now notes that the
  leading `=` must be omitted; the engine adds it internally.

## [0.10.0] - 2026-03-28

### Added

- Artifact-emitted discovery commands: `--print-request-template` for a minimal valid request and
  `--print-protocol-catalog` for the machine-readable protocol inventory.
- Protocol-owned catalog metadata covering source types, persistence types, operations, reads, and
  nested tagged request unions.
- Request-shape error classification via the new `INVALID_REQUEST_SHAPE` problem code.

### Changed

- All tagged request unions now use `type` as the discriminator field, including `source`,
  `persistence`, selections, and success-side persistence outcomes.
- CLI help is now first-success capable: it prints the current version, a minimal valid request,
  stdin and Docker-mounted examples, discovery commands, Docker path semantics, and version-routed
  public doc links.
- Failure contexts now surface `sourceType` and `persistenceType` fields, matching the request
  contract terminology.
- Public docs, runnable examples, and committed Jazzer request seeds now use the unified `type`
  contract and describe the new discovery and error-taxonomy behavior.
- Jazzer protocol-request replay and assertions now distinguish syntax failures from request-shape
  failures and semantic validation failures.

### Fixed

- Syntactically valid payloads with missing required fields, unknown discriminator IDs, or wrong
  token shapes are no longer mislabeled as `INVALID_JSON`.

## [0.9.0] - 2026-03-27

### Added

- Finding-bearing document analysis reads: `ANALYZE_FORMULA_HEALTH`,
  `ANALYZE_HYPERLINK_HEALTH`, `ANALYZE_NAMED_RANGE_HEALTH`, and
  `ANALYZE_WORKBOOK_FINDINGS`.
- Real CLI help output via `--help` and `-h`, including usage guidance for stdin/stdout,
  request/response files, Docker, and the packaged Java 26 fat JAR.

### Changed

- Factual read operations now use `GET_*` names consistently: formula surface, sheet schema, and
  named-range surface are exposed as introspection reads instead of conclusion-bearing
  `ANALYZE_*_SURFACE` operations.
- Success persistence outcomes now distinguish `NOT_SAVED`, `SAVED_AS`, and `OVERWRITTEN`, with
  explicit caller-facing path tokens plus execution-environment paths for saved workbooks.
- Internal document-intelligence architecture now uses a shared analysis-finding model across
  formula, hyperlink, named-range, and aggregate workbook health reads.
- Jazzer generators, assertions, readable request seeds, and live-run invariants now enforce the
  renamed read taxonomy and the new persistence/analysis contracts.
- Refreshed pinned GitHub Actions workflow dependencies to current Node 24-ready releases, replaced the release-publish action with a GitHub CLI release step, and configured Dependabot to stop reopening the rejected `gradle/actions` v6 major upgrade.

## [0.8.0] - 2026-03-27

### Changed
- The CLI and protocol architecture now expose a transport-neutral request-executor port and a
  thinner CLI transport boundary, keeping argument parsing, protocol I/O, and execution concerns
  more cleanly separated.
- Public product wording now describes GridGrind as agent-first but not agent-only, matching the
  shipped JSON protocol, CLI/container transport, and current non-MCP distribution model.
- Local Jazzer regression verification now replays committed seeds through four isolated
  per-harness launcher tasks before producing the aggregate regression summary, and local Jazzer
  harness execution no longer depends on Gradle's flaky binary test-results pipeline.
- The request and success protocol now use ordered `reads` and `persistence` outcomes instead of
  the old aggregated `analysis` and nullable saved-path model, making post-mutation workbook
  introspection and insights explicit, typed, and request-correlated.

### Added
- Read operations for workbook summary, named ranges, sheet summary, cells, windows,
  merged regions, hyperlinks, comments, sheet layout, formula surface, sheet schema, and
  named-range surface.
- Public read-heavy example request in `examples/introspection-analysis-request.json`.
- Jazzer request-seed coverage for the new read-heavy public example and the broader `reads`
  protocol shape.

## [0.7.0] - 2026-03-27

### Added
- Cell-metadata authoring with `SET_HYPERLINK`, `CLEAR_HYPERLINK`, `SET_COMMENT`, and
  `CLEAR_COMMENT`, plus typed hyperlink and comment analysis in cell reports and previews.
- Workbook named-range authoring with `SET_NAMED_RANGE` and `DELETE_NAMED_RANGE`, plus requested
  workbook-level named-range analysis and named-range counts in workbook summaries.
- Public example request for hyperlink, comment, and named-range workflows in
  `examples/excel-authoring-essentials-request.json`.
- Jazzer regression seeds covering the public authoring example, hyperlink/comment workflow
  failures, named-range workflow failures, named-range normalization round-trips, and additional
  authoring-metadata replay cases.

### Changed
- Named-range targets are now canonicalized to top-left:`bottom-right` order on input, so shapes
  such as `B2:A1` are stored and reported as `A1:B2`.
- `CLEAR_RANGE` now documents its full effect on cell metadata: it clears hyperlink and comment
  state in addition to values and styles.
- Jazzer coverage and operator docs now describe the expanded authoring surface and the larger
  committed seed floor.

### Fixed
- Repeated `SET_HYPERLINK` writes on the same cell now preserve the latest hyperlink target after
  `.xlsx` save and reopen instead of leaking an older target through the persisted workbook.
- `.xlsx` round-trip verification no longer treats named-range target normalization as a failure
  after save and reopen.

## [0.6.0] - 2026-03-26

### Added
- Local-only Jazzer fuzzing layer in a separate nested Gradle build under `jazzer/`, including
  protocol request fuzzing, structured workflow fuzzing, engine command-sequence fuzzing, and
  `.xlsx` round-trip fuzzing.
- Convenience wrapper scripts under `jazzer/bin/` for regression replay, per-harness fuzzing,
  aggregate fuzzing, and local cleanup.
- Task-specific local corpus storage under `jazzer/.local/runs/*/.cifuzz-corpus/`.
- Jazzer operator commands for latest-summary status/report views, corpus inspection, finding
  listing, one-off input replay, and seed promotion.
- Per-target Jazzer run history with latest-summary JSON/text artifacts, per-harness telemetry,
  and replayed local finding artifacts.
- Committed custom seed floor for all four current Jazzer harnesses, including readable public
  request examples and replay-verified binary workflow seeds, plus promotion metadata under
  `jazzer/src/fuzz/resources/dev/erst/gridgrind/jazzer/promoted-metadata/`.
- Deterministic Jazzer support tests for summary parsing, summary rendering, and `.xlsx`
  round-trip verifier behavior inside the nested `jazzer/` build.
- Richer `.xlsx` style authoring through `APPLY_STYLE`, including `fontName`,
  typed `fontHeight`, `fontColor`, `underline`, `strikeout`, `fillColor`, and side-aware border
  patches with `all` defaults plus top/right/bottom/left overrides.
- Effective style analysis and `.xlsx` round-trip verification for font, fill, and border facts,
  plus a dedicated formatting-depth example request.

### Changed
- Developer documentation now includes an authoritative Jazzer architecture and operations
  reference in `docs/DEVELOPER_JAZZER.md`.
- The Jazzer workflow now uses lock-protected wrapper scripts so only one local Jazzer command
  runs at a time, avoiding concurrent runtime-initialization failures.
- Jazzer documentation is now split into architecture, operations, and coverage references so the
  local fuzzing layer, command surface, and current harness inventory can be read independently.
- `jazzer/bin/fuzz-all` now runs the four active harness scripts sequentially so every harness
  still gets its own lock, run history, summary, and telemetry artifacts.
- Jazzer workflow fuzzing now exercises `NEW` and `EXISTING` source modes together with `NONE`,
  `SAVE_AS`, and `OVERWRITE` persistence modes instead of staying on new-workbook, no-persistence
  flows only.
- Root project-file formatting now excludes local-only instruction and scratch directories, keeping
  local workspace state outside the canonical quality gates.
- `jazzer/bin/list-corpus` now reports generated local corpus and committed custom seeds
  separately, making the active-fuzz seed floor easier to interpret.
- The committed Jazzer regression seed floor now includes the formerly crashing malformed-formula
  engine command sequence as an expected-invalid replay case.
- Public operation and quick-reference docs now describe the expanded style contract, including
  RGB color normalization, solid-fill semantics, border-style enums, and analyzed effective style
  output.
- The public style contract now uses typed `fontHeight` objects instead of the old integer
  `fontSizePoints` field. Requests can express font height as exact points or exact twips, and
  analyzed style reports now return both `fontHeight.points` and `fontHeight.twips`.

### Fixed
- Malformed formulas that trigger Apache POI parser-state `IllegalStateException`s are now
  surfaced as `INVALID_FORMULA` instead of leaking as `INTERNAL_ERROR`.
- Jazzer `.xlsx` round-trip fuzzing now validates persisted formatting-depth style state
  accurately instead of relying on coarse style heuristics, so alignment-only and exact
  `fontHeight` round-trips no longer produce false positives.
- Jazzer latest-summary status and report views now distinguish active findings from expected-
  invalid and replay-clean local artifacts, preventing stale local crash files from being
  misreported as current failures.
- Jazzer summary parsing now handles active-fuzz corpus-size output that uses `Kb` units, keeping
  latest-summary metrics accurate for longer local fuzzing runs.

## [0.5.0] - 2026-03-25

### Added
- Explicit `.xlsx`-only request contract for existing-workbook and save-as paths, with deterministic rejection of `.xls`, `.xlsm`, `.xlsb`, and other non-`.xlsx` workbook paths.
- Reusable `.xlsx` round-trip test infrastructure for reopening generated workbooks and asserting sheet order, merged regions, column widths, row heights, and freeze-pane state.
- Sheet-management operations: `RENAME_SHEET`, `DELETE_SHEET`, and `MOVE_SHEET`.
- Structural layout operations: `MERGE_CELLS`, `UNMERGE_CELLS`, `SET_COLUMN_WIDTH`, `SET_ROW_HEIGHT`, and `FREEZE_PANES`.
- Runnable example requests for sheet management and structural layout workflows.
- Expanded capability documentation describing current Excel support, shipped behavior, and remaining Apache POI parity gaps.

### Changed
- GridGrind workbook support is now an explicit product contract for `XSSF .xlsx` only rather than an inferred capability.
- Public operation and error documentation now covers strict sheet-management and structural-layout semantics, including exact-match unmerge behavior, explicit width and height units, and freeze-pane coordinate rules.
- Developer documentation now includes the expanded runnable example inventory.
- Project documentation now summarizes the shipped `.xlsx`, sheet-management, and structural-layout scope together with verification status and the next recommended parity area.

### Fixed
- Capability inventory contradictions around merged cells and freeze panes removed so the standing agent inventory matches the shipped protocol surface.
- Structural layout coverage gaps in protocol and engine tests closed, including direct helper-branch verification for merge overlap checks, width and height validation, exact merged-region lookup, and freeze-pane coordinate validation.

## [0.4.1] - 2026-03-25

### Added
- `--version` flag: prints `gridgrind <version>` to stdout and exits with code 0.

## [0.4.0] - 2026-03-25

### Security
- All GitHub Actions action references pinned to exact commit SHAs; Dependabot configured to
  keep pins current automatically.

### Changed
- CI workflow now runs `./gradlew check` only; redundant `coverage` task removed from CI since
  report generation without artifact upload produced files that were immediately discarded.
- Job timeouts added to all three workflows: 15 minutes for CI, 20 minutes for Release and
  Container.
- `FormulaException` converted from abstract sealed class to sealed interface; `InvalidFormulaException`
  and `UnsupportedFormulaException` now extend `IllegalArgumentException` directly and implement the
  interface, carrying their own fields.
- `PayloadException` converted from abstract sealed class to sealed interface; `InvalidJsonException`
  and `InvalidRequestException` now extend `IllegalArgumentException` directly and implement the
  interface, carrying their own fields.
- `ExcelRange` converted from a manually managed class to a record with a compact constructor that
  enforces non-negative and ordered bounds.
- `ExcelCellStyleSnapshot` alignment fields changed from `String` to typed `ExcelHorizontalAlignment`
  and `ExcelVerticalAlignment` enums.
- `CellStyleInput` shadow enums `HorizontalAlignmentInput` and `VerticalAlignmentInput` removed;
  fields now use `ExcelHorizontalAlignment` and `ExcelVerticalAlignment` directly from the engine.
- `ExcelPreviewRow` compact constructor added to enforce an unmodifiable defensive copy of the cells
  list and coerce null to an empty list.
- `GridGrindProblems.enrichContext` exhaustively enumerates all nine `ProblemContext` subtypes
  without a `default` arm; compiler now enforces exhaustiveness.
- `GridGrindService.closeWorkbook` converted from `instanceof` conditional to exhaustive pattern-
  matching switch over the sealed `GridGrindResponse` type.
- `ExcelSheet.lastColumnIndex` and `ExcelSheet.save` dead null-check branches removed; POI
  iterators never yield null rows and `Path.toAbsolutePath().getParent()` is always non-null.
- `ExcelRange.parseCell` dead null-check on the address parameter removed.
- `WorkbookStyleRegistry.resolveNumberFormat` extracted as a package-private static method to
  cover the null/blank substitution branch directly without reflection.
- `GridGrindJson.jsonLine` and `GridGrindJson.jsonColumn` widened to package-private to allow
  direct testing of the null-location and non-positive-line-number branches.
- `GridGrindProblems.enrichContext` widened to package-private to allow direct testing of context
  types that are never paired with enrichable exceptions in integration paths.
- `GridGrindService.formulaFor`, `sheetNameFor`, `addressFor`, `rangeFor`, and
  `GridGrindProblems.enrichContext` multi-label pattern case arms split into individual per-subtype
  arms, and `when` guards replaced with if-else blocks inside the relevant arms, eliminating
  JaCoCo false-positive missed branches that Java 26 pattern switches generate for multi-label
  case arms.
- `GridGrindService.validateRequest` `None`/`SaveAs` multi-label arm split into two individual
  arms for the same reason.
- `GridGrindJson.jsonPath` dead `if (index >= 0)` guard removed; `getIndex()` always returns
  a non-negative value for array-position references.
- `ExcelSheet.snapshot` formula expression extracted before the try block so the catch arm can
  reference it without a ternary inside the catch, removing an unreachable branch.
- `GridGrindService.formulaFor` `when op.value() instanceof CellInput.Formula` guard replaced
  with an inner exhaustive switch over the sealed `CellInput` type, eliminating the guard-false
  synthetic branch that JaCoCo counted as a missed branch.

## [0.3.0] - 2026-03-25

### Changed
- `source.mode` wire value renamed: `EXISTING_FILE` is now `EXISTING`.
- `persistence.mode` wire value renamed: `OVERWRITE_SOURCE` is now `OVERWRITE`. The `OVERWRITE`
  mode no longer accepts a `path` field; it always overwrites the source file.
- `AUTO_SIZE_COLUMNS` operation no longer accepts a `columns` field; it always auto-sizes all
  populated columns on the sheet.
- `CellInput.Date` component field renamed from `localDate` to `date`; `CellInput.DateTime`
  component field renamed from `localDateTime` to `dateTime`.

## [0.2.0] - 2026-03-25

### Added
- Native `linux/arm64` container image published alongside `linux/amd64`. Apple Silicon Macs,
  ARM Linux, and Windows ARM pull the correct image automatically with no `--platform` flag.

### Fixed
- Error reference documentation corrected: category values, recovery strategy names, problem
  code table, and causes chain fields now match the actual wire protocol.

## [0.1.0] - 2026-03-24

### Added
- Initial release.

[Unreleased]: https://github.com/resoltico/GridGrind/compare/v0.60.0...HEAD
[0.60.0]: https://github.com/resoltico/GridGrind/compare/v0.59.0...v0.60.0
[0.59.0]: https://github.com/resoltico/GridGrind/compare/v0.58.0...v0.59.0
[0.58.0]: https://github.com/resoltico/GridGrind/compare/v0.57.0...v0.58.0
[0.57.0]: https://github.com/resoltico/GridGrind/compare/v0.56.0...v0.57.0
[0.56.0]: https://github.com/resoltico/GridGrind/compare/v0.55.0...v0.56.0
[0.55.0]: https://github.com/resoltico/GridGrind/compare/v0.54.0...v0.55.0
[0.54.0]: https://github.com/resoltico/GridGrind/compare/v0.53.0...v0.54.0
[0.53.0]: https://github.com/resoltico/GridGrind/compare/v0.52.0...v0.53.0
[0.52.0]: https://github.com/resoltico/GridGrind/compare/v0.51.0...v0.52.0
[0.51.0]: https://github.com/resoltico/GridGrind/compare/v0.50.0...v0.51.0
[0.50.0]: https://github.com/resoltico/GridGrind/compare/v0.49.0...v0.50.0
[0.49.0]: https://github.com/resoltico/GridGrind/compare/v0.48.0...v0.49.0
[0.48.0]: https://github.com/resoltico/GridGrind/compare/v0.47.0...v0.48.0
[0.47.0]: https://github.com/resoltico/GridGrind/compare/v0.46.0...v0.47.0
[0.46.0]: https://github.com/resoltico/GridGrind/compare/v0.45.0...v0.46.0
[0.45.0]: https://github.com/resoltico/GridGrind/compare/v0.44.0...v0.45.0
[0.44.0]: https://github.com/resoltico/GridGrind/compare/v0.43.0...v0.44.0
[0.43.0]: https://github.com/resoltico/GridGrind/compare/v0.42.0...v0.43.0
[0.42.0]: https://github.com/resoltico/GridGrind/compare/v0.41.0...v0.42.0
[0.41.0]: https://github.com/resoltico/GridGrind/compare/v0.40.0...v0.41.0
[0.40.0]: https://github.com/resoltico/GridGrind/compare/v0.39.0...v0.40.0
[0.39.0]: https://github.com/resoltico/GridGrind/compare/v0.38.0...v0.39.0
[0.38.0]: https://github.com/resoltico/GridGrind/compare/v0.37.0...v0.38.0
[0.37.0]: https://github.com/resoltico/GridGrind/compare/v0.36.0...v0.37.0
[0.36.0]: https://github.com/resoltico/GridGrind/compare/v0.35.0...v0.36.0
[0.35.0]: https://github.com/resoltico/GridGrind/compare/v0.34.0...v0.35.0
[0.34.0]: https://github.com/resoltico/GridGrind/compare/v0.33.0...v0.34.0
[0.33.0]: https://github.com/resoltico/GridGrind/compare/v0.32.2...v0.33.0
[0.32.2]: https://github.com/resoltico/GridGrind/compare/v0.32.1...v0.32.2
[0.32.1]: https://github.com/resoltico/GridGrind/compare/v0.32.0...v0.32.1
[0.32.0]: https://github.com/resoltico/GridGrind/compare/v0.31.0...v0.32.0
[0.31.0]: https://github.com/resoltico/GridGrind/compare/v0.30.0...v0.31.0
[0.30.0]: https://github.com/resoltico/GridGrind/compare/v0.29.0...v0.30.0
[0.29.0]: https://github.com/resoltico/GridGrind/compare/v0.28.0...v0.29.0
[0.28.0]: https://github.com/resoltico/GridGrind/compare/v0.27.0...v0.28.0
[0.27.0]: https://github.com/resoltico/GridGrind/compare/v0.26.0...v0.27.0
[0.26.0]: https://github.com/resoltico/GridGrind/compare/v0.25.0...v0.26.0
[0.25.0]: https://github.com/resoltico/GridGrind/compare/v0.24.0...v0.25.0
[0.24.0]: https://github.com/resoltico/GridGrind/compare/v0.23.0...v0.24.0
[0.23.0]: https://github.com/resoltico/GridGrind/compare/v0.22.0...v0.23.0
[0.22.0]: https://github.com/resoltico/GridGrind/compare/v0.21.0...v0.22.0
[0.21.0]: https://github.com/resoltico/GridGrind/compare/v0.20.0...v0.21.0
[0.20.0]: https://github.com/resoltico/GridGrind/compare/v0.19.0...v0.20.0
[0.19.0]: https://github.com/resoltico/GridGrind/compare/v0.18.0...v0.19.0
[0.18.0]: https://github.com/resoltico/GridGrind/compare/v0.17.0...v0.18.0
[0.17.0]: https://github.com/resoltico/GridGrind/compare/v0.16.0...v0.17.0
[0.16.0]: https://github.com/resoltico/GridGrind/compare/v0.15.0...v0.16.0
[0.15.0]: https://github.com/resoltico/GridGrind/compare/v0.14.0...v0.15.0
[0.14.0]: https://github.com/resoltico/GridGrind/compare/v0.13.0...v0.14.0
[0.13.0]: https://github.com/resoltico/GridGrind/compare/v0.12.0...v0.13.0
[0.12.0]: https://github.com/resoltico/GridGrind/compare/v0.11.0...v0.12.0
[0.11.0]: https://github.com/resoltico/GridGrind/compare/v0.10.0...v0.11.0
[0.10.0]: https://github.com/resoltico/GridGrind/compare/v0.9.0...v0.10.0
[0.9.0]: https://github.com/resoltico/GridGrind/compare/v0.8.0...v0.9.0
[0.8.0]: https://github.com/resoltico/GridGrind/compare/v0.7.0...v0.8.0
[0.7.0]: https://github.com/resoltico/GridGrind/compare/v0.6.0...v0.7.0
[0.6.0]: https://github.com/resoltico/GridGrind/compare/v0.5.0...v0.6.0
[0.5.0]: https://github.com/resoltico/GridGrind/compare/v0.4.1...v0.5.0
[0.4.1]: https://github.com/resoltico/GridGrind/compare/v0.4.0...v0.4.1
[0.4.0]: https://github.com/resoltico/GridGrind/compare/v0.3.0...v0.4.0
[0.3.0]: https://github.com/resoltico/GridGrind/compare/v0.2.0...v0.3.0
[0.2.0]: https://github.com/resoltico/GridGrind/compare/v0.1.0...v0.2.0
[0.1.0]: https://github.com/resoltico/GridGrind/releases/tag/v0.1.0
