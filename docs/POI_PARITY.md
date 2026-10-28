# Apache POI parity

GridGrind uses Apache POI 6.0.0 for its workbook backend. It has a scoped behavioral parity
suite for the modeled `.xlsx` contract. It does not claim complete coverage of the POI API,
every OOXML structure, or Microsoft Excel. Macros are one exclusion, not the sole exclusion.

## What the evidence establishes

The [executable ledger](../engine/src/parityTest/resources/dev/erst/gridgrind/engine/runtime/parity/xssf-parity-ledger.json)
is the authoritative row inventory. Its 71 rows declare COMPLETE and expect PASS, backed by
47 distinct probes and a [19-scenario corpus](../engine/src/parityTest/resources/dev/erst/gridgrind/engine/runtime/parity/xssf-parity-corpus.json).
Several rows share a probe. These figures count defined checks, not a percentage of POI or
Excel functionality.

[XlsxParityTest](../engine/src/parityTest/java/dev/erst/gridgrind/engine/runtime/parity/XlsxParityTest.java)
checks that ledger/corpus versions match the actual POI runtime, materializes the fixtures,
and compares every row's expected outcome with its observed probe result. The probes compare
GridGrind operations and reopened workbook facts with POI/OOXML comparison code. Signature
verification reuses the production XML-DSig provider factory, so that path is not independent
evidence for provider configuration. COMPLETE describes the declared row's scope; it is not an independent inventory of every POI capability.

Run the focused suite with:

    ./gradlew :engine:parityTest

A passing run establishes agreement for those scenarios under the tested runtime. It does
not establish native Excel rendering, recalculation, VBA execution, or exhaustive input coverage.

## Capability matrix

This summarizes the tested families. Exact shapes come from the [protocol reference](reference/README.md);
the ledger and probe implementations own the detailed evidence.

| Area | Tested scope | Boundary |
| --- | --- | --- |
| Workbook and sheets | Summary facts, lifecycle, protection, copy, merged regions, layout and printing | Modeled properties and declared fixtures; structural restrictions remain in [Limits](LIMITATIONS.md). |
| Cells and styles | Typed/rich text, modeled style properties, theme/indexed colors, tints and gradients | Factual readback and authored changes have separate contracts; no visual-equivalence certification. |
| Comments, links and names | Plain/rich comments and anchors, hyperlinks, explicit and formula-defined names | Authored identifier and formula-security constraints apply. |
| Tables, filters and validation | Modeled metadata, criteria, sort state, validation and conditional-format families | Unsupported native detail can be reported or preserved without becoming authorable. |
| Formulas | Core evaluation, external workbooks, missing-workbook policy, template UDFs, caches and targeted evaluation | POI evaluation and the declared environment; not every Excel function or add-in. |
| Drawings and embedded objects | Pictures, shapes, connectors, anchors, embedded payload readback and preservation | The modeled drawing contract; preservation does not imply arbitrary OLE execution or editing. |
| Charts | BAR/LINE/PIE oracle projections, authoring/replacement fixtures and BAR/LINE combo handling | The product models additional plot families; this ledger does not exhaustively compare them or every chart option/renderer. |
| Pivot tables | Readback/health, range/name/table-backed authoring, replacement/deletion and preservation | The modeled pivot contract; no certification of every Excel pivot feature or refresh result. |
| Large-sheet modes | Event-reader summaries and append-oriented streaming authoring | EVENT_READ supports workbook/sheet summaries but existing-source preflight uses full XSSF; STREAMING_WRITE is restricted NEW-workbook authoring and can materialize XSSF for inspection, assertions and publication checks. See [Execution model](EXECUTION_MODEL.md). |
| Package security | Password handling, strong AGILE encryption, signature inspection, unsigned output and signing | Explicit grants and security dispositions; broader observed encryption state is not all authorable. |
| Other file formats and VBA | Outside the GridGrind product contract | `.xls`/`.xlsb` content and `.xlsm` filenames are rejected; there is no VBA authoring or execution interface. |

The `.xlsx` contract does not certify that a renamed or externally authored OOXML package is
free of VBA or other active content. File-format admission is not a macro-stripping service.

## Interpreting POI support

POI's [spreadsheet overview](https://poi.apache.org/components/spreadsheet/) covers APIs broader
than GridGrind's `.xlsx` interface, including HSSF and lower-level access. Its
[limitations](https://poi.apache.org/components/spreadsheet/limitations.html) also distinguish
limited chart/pivot support and macro preservation from macro creation. Those upstream pages
provide context; they are not an exhaustive POI 6 feature inventory or GridGrind acceptance test.

POI 6's [release notes](https://poi.apache.org/changes.html) include behavior changes in formula
calculation, formatting and drawing shifts. GridGrind has dedicated regressions for unary-minus
precedence, persisted multiplication/division precision and picture-anchor movement in
[formula tests](../engine/src/test/java/dev/erst/gridgrind/engine/backend/xssf/ExcelWorkbookFormulasTest.java)
and [structural-edit tests](../engine/src/test/java/dev/erst/gridgrind/engine/backend/xssf/ExcelRowStructureEditCoverageTest.java).
They complement the ledger; they do not expand it into an exhaustive comparison with all POI 6 changes.
