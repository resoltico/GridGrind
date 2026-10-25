# Golden behaviour corpus

The corpus runs the current public CLI in process and through the packaged shadow
JAR. Both launchers must match the same snapshots; the packaged launcher never
updates them. No product code was refactored to create this baseline.

## Coverage and observations

- Every checked-in `examples/*.json` file is discovered from disk and executed
  with its declared recipe assets and an explicit fixture grant.
- Every registry recipe is printed, materialized into an isolated workspace,
  looked up in the discovery catalog, and executed. Asset-backed print rejection
  is captured as part of the current contract.
- Every discovery command family is exercised, including the three help topics,
  text and structured identity output, templates, catalog search and lookup,
  recipe catalog and keyword matching.
- Doctor inputs cover invalid UTF-8, duplicate keys and step ids, unknown fields,
  required-field omission, explicit null, malformed scalars, missing/unknown
  discriminators, constructor validation, invalid enums, out-of-bounds addresses,
  multiple independent defects, truncation, and valid-sibling syntax recovery.
  These cases follow the intake claims in README and
  `docs/REQUEST_AND_EXECUTION_REFERENCE.md`, the problem families in
  `docs/ERRORS.md`, and the existing CLI encoding and doctor tests.
  Raw `.bin` inputs are marked binary in `.gitattributes` to preserve invalid
  UTF-8 and the exact fuzz input, including its trailing whitespace.
- Execution failures cover missing workbook, sheet, and source-backed input.
- Every produced workbook is reopened through GridGrind. Summary and per-sheet
  bounds determine exhaustive cell addresses for `GET_CELLS`; readback includes
  every cell facet, sheet and print layout, tables, pivots, charts, drawings,
  validation, conditional formatting, and workbook findings. There is no POI
  oracle or binary workbook comparison in the golden tests.

The one-character negative control changes a `SET_CELL` operand from `alpha` to
`alphb`. Both launchers reopen the workbook, verify that `alphb` was observed,
and require snapshot comparison to fail. This control never updates its mutated
observation into the baseline.

## Normalization

Object keys are sorted; array ordering is retained. Workspace and repository
paths, lifecycle timestamps, durations, and implementation versions are masked.
Published workbook SHA-256 and byte size vary with ZIP metadata and generated
OOXML identifiers, so those artifact identities are masked. Source-input and
preview identities are retained. Workbook facts are compared directly.

Signature-line setup UUIDs are generated on authoring. Their syntax is checked,
then each distinct UUID receives a deterministic ordinal within the observation.
The same UUID must therefore remain consistent across mutation reads and reopened
workbook facts. UUID normalization does not mask cell strings or other identifiers.

## Recording and verification

From the repository root:

```sh
./gradlew :cli:test --tests '*GoldenCorpusTest' -Pgolden.update=true --no-parallel
./gradlew :cli:test --tests '*GoldenCorpusTest' --no-parallel
./gradlew check --continue --no-parallel
git diff -- cli/src/test/resources/golden/
```

`cli:test` depends on `cli:shadowJar`. Update mode writes only in-process
observations to `cli/src/test/resources/golden/snapshots/`. The packaged launcher
must still compare successfully. Review every changed snapshot and explain each
intentional wire or behaviour change in the same commit message. An ordinary run
fails on missing snapshots; it never silently blesses new cases.

## Known baseline defect

S2's active fuzz run found a malformed object key that makes diagnostic
construction throw `jsonPath must not be blank`. The exact input is retained as
`problems/blank-key-syntax.bin`; the doctor snapshot records the current
`INVALID_REQUEST` report with unavailable JSON location. This is a known defect,
not the required outcome. The intake repair must replace it with classified
syntax findings and preserve valid siblings; its snapshot diff must be explained.

## Initial harness failures

The first recording attempt ran 139 tests and failed 33: 30 asserted the wrong
doctor exit code, and 3 exposed generated signature UUIDs. The second ran 141
tests (including the retained fuzz input) and failed 30 because the doctor
expectation was incorrectly changed to 0. `CliExitCodes.forDoctorReport` and the
JUnit failure outputs establish the correct value, 1. These were harness errors;
the product code was unchanged. Full logs and the first run's XML are retained
outside the repository under the package's `evidence/wp-01/` directory.

## Verified initial corpus

The final no-update run discovers 141 golden tests: both launchers for 36 example
and recipe executions, 15 discovery cases, 16 doctor cases and 3 execution
failures, plus the negative-control test that exercises both launchers. All 141
passed with 0 errors and 0 skips (319.122 s). There are 71 shared snapshots; no
generated workbook is stored in the corpus.

The full `./gradlew check --continue --no-parallel` run took 370.87 s. Its 2821
tests had only the 3 approved pre-existing policy-freshness failures; all 596 CLI
tests passed. Formatting and PMD passed. The only failing tasks were the three
expired shape-policy gates and `build-logic:test`, as allowed for Phase 0.
`git diff main -- AGENTS.md` remains empty.

This temporary record is removed in WP-80 after its evidence is carried into the
final verification report.
