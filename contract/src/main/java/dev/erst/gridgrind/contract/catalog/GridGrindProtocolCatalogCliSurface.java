package dev.erst.gridgrind.contract.catalog;

import java.util.List;

/** Owns the public CLI/help surface published alongside the protocol catalog. */
final class GridGrindProtocolCatalogCliSurface {
  static final CliSurface CLI_SURFACE =
      new CliSurface(
          new CliSurface.CliSection(
              "Usage",
              List.of(
                  "gridgrind [--request <path>] [--response <path>]",
                  "gridgrind --doctor-request [--request <path>]",
                  "gridgrind --print-request-template",
                  "gridgrind --print-task-catalog [--task <id>]",
                  "gridgrind --print-task-plan <id>",
                  "gridgrind --print-goal-plan <goal>",
                  "gridgrind --print-protocol-catalog",
                  "gridgrind --print-protocol-catalog --operation <id>|<group>:<id>",
                  "gridgrind --print-protocol-catalog --search <text>",
                  "gridgrind --print-example <id>",
                  "gridgrind --help | -h | help",
                  "gridgrind --version")),
          new CliSurface.CliSection(
              "Execution",
              List.of(
                  "GridGrind executes ordered steps in sequence, then saves the workbook (unless"
                      + " persistence is NONE); if any step fails, no workbook is written.",
                  "A NEW workbook starts with zero sheets; use ENSURE_SHEET to create the first"
                      + " sheet.",
                  "execution is optional; omit it for the default FULL_XSSF read and write path"
                      + " plus NORMAL journaling.")),
          new CliSurface.CliDefinitionSection(
              "Limits",
              List.of(
                  new CliSurface.DefinitionEntry(
                      "File format", ".xlsx only; .xls, .xlsm, and .xlsb are rejected."),
                  new CliSurface.DefinitionEntry(
                      "Sheet names",
                      "1 to 31 characters; reject : \\ / ? * [ ] and leading/trailing"
                          + " apostrophes."),
                  new CliSurface.DefinitionEntry(
                      "GET_WINDOW cell count", "rowCount * columnCount must not exceed 250,000."),
                  new CliSurface.DefinitionEntry(
                      "GET_SHEET_SCHEMA cells", "rowCount * columnCount must not exceed 250,000."),
                  new CliSurface.DefinitionEntry(
                      "Request JSON size", GridGrindContractText.requestDocumentLimitSummary()),
                  new CliSurface.DefinitionEntry(
                      "EVENT_READ mode",
                      GridGrindExecutionModeMetadata.eventRead().catalogSummary()),
                  new CliSurface.DefinitionEntry(
                      "STREAMING_WRITE mode",
                      GridGrindExecutionModeMetadata.streamingWrite().catalogSummary()),
                  new CliSurface.DefinitionEntry(
                      "Column widthCharacters",
                      "authored widthCharacters > 0 and <= 255 (Excel limit)."),
                  new CliSurface.DefinitionEntry(
                      "Default sheet sizing",
                      "authored defaultColumnWidth must be > 0 and <= 255;"
                          + " authored defaultRowHeightPoints must be > 0 and <= "
                          + dev.erst.gridgrind.excel.foundation.ExcelSheetLayoutLimits
                              .MAX_ROW_HEIGHT_POINTS
                          + " (Excel limits)."),
                  new CliSurface.DefinitionEntry(
                      "Row heightPoints",
                      "authored heightPoints > 0 and <= "
                          + dev.erst.gridgrind.excel.foundation.ExcelSheetLayoutLimits
                              .MAX_ROW_HEIGHT_POINTS
                          + " (Excel row height limit)."),
                  new CliSurface.DefinitionEntry(
                      "Zoom percent",
                      dev.erst.gridgrind.excel.foundation.ExcelSheetLayoutLimits.MIN_ZOOM_PERCENT
                          + " to "
                          + dev.erst.gridgrind.excel.foundation.ExcelSheetLayoutLimits
                              .MAX_ZOOM_PERCENT
                          + " inclusive (Excel zoom limit)."),
                  new CliSurface.DefinitionEntry(
                      "Row structural edits",
                      "rejected when they would move tables, sheet autofilters, or data"
                          + " validations; deletes/shifts also reject destructive range-backed"
                          + " named ranges."),
                  new CliSurface.DefinitionEntry(
                      "Column structural edits",
                      "same ownership rule; deletes/shifts also reject destructive range-backed"
                          + " named ranges; all column edits reject any workbook formulas or"
                          + " formula-defined names."),
                  new CliSurface.DefinitionEntry(
                      "Chart mutations",
                      "SET_CHART authors AREA, AREA_3D, BAR, BAR_3D, DOUGHNUT, LINE,"
                          + " LINE_3D, PIE, PIE_3D, RADAR, SCATTER, SURFACE, and SURFACE_3D;"
                          + " unsupported loaded chart detail is preserved on unrelated edits"
                          + " and rejected for"
                          + " authoritative mutation."),
                  new CliSurface.DefinitionEntry(
                      "Chart title formulas",
                      "SET_CHART title FORMULA and series.title FORMULA must resolve to one"
                          + " cell, directly or through one defined name."),
                  new CliSurface.DefinitionEntry(
                      "Array formulas",
                      "SET_ARRAY_FORMULA authors one contiguous single-cell or multi-cell array"
                          + " formula group; CLEAR_ARRAY_FORMULA may target any member cell and"
                          + " removes the whole stored group."),
                  new CliSurface.DefinitionEntry(
                      "Drawing validation",
                      "failed SET_SHAPE / SET_CHART validation leaves existing drawing state"
                          + " unchanged and creates no partial artifacts."),
                  new CliSurface.DefinitionEntry(
                      "DATE / DATE_TIME inputs",
                      "stored as numeric serial; GET_CELLS returns declaredType=NUMBER."),
                  new CliSurface.DefinitionEntry(
                      "Formula authoring", GridGrindContractText.formulaAuthoringLimitSummary()),
                  new CliSurface.DefinitionEntry(
                      "Loaded formula support",
                      GridGrindContractText.loadedFormulaSupportSummary()))),
          new CliSurface.CliSection(
              "Request",
              List.of(
                  "protocolVersion is optional; omit it and the current version is assumed.",
                  "persistence is optional; omit it and the workbook stays in memory only"
                      + " (NONE).",
                  "planId is optional; omit it and GridGrind will generate one for the"
                      + " execution journal.",
                  "execution is optional; omit it for FULL_XSSF reads and writes with NORMAL"
                      + " journaling, or supply execution.mode to choose EVENT_READ or"
                      + " STREAMING_WRITE when their limits fit the plan.",
                  "execution.journal.level controls journal detail; VERBOSE also streams live"
                      + " phase events to stderr.",
                  "execution.calculation controls server-side evaluation, cache clearing, and"
                      + " open-time recalc flags. "
                      + GridGrindContractText.calculationPolicyInputSummary(),
                  "Source-backed text and binary fields support INLINE, UTF8_FILE or FILE, and"
                      + " STANDARD_INPUT sources.",
                  "Large authored payloads belong in UTF8_FILE, FILE, or STANDARD_INPUT"
                      + " sources; the request JSON itself is capped at 16 MiB.",
                  GridGrindContractText.standardInputRequiresRequestMessage()
                      + " because stdin cannot carry both the request JSON and authored input"
                      + " content in one CLI invocation.",
                  "formulaEnvironment is optional; omit it for the default evaluator, or supply"
                      + " it to bind external workbooks, choose missing-workbook policy, and"
                      + " register template-backed UDFs.",
                  "steps is optional; omit or send [] for an empty no-op plan.",
                  GridGrindContractText.stepKindSummary(),
                  "Step order is authoritative: mutations, assertions, and inspections may be"
                      + " interleaved when the workflow demands it.")),
          new CliSurface.CliDefinitionSection(
              "File Workflow",
              List.of(
                  new CliSurface.DefinitionEntry(
                      "No --request flag", "read the JSON request from stdin."),
                  new CliSurface.DefinitionEntry(
                      "--request <path>", "read the JSON request from that file."),
                  new CliSurface.DefinitionEntry(
                      "No --response flag", "write the JSON response to stdout."),
                  new CliSurface.DefinitionEntry(
                      "--response <path>",
                      "write the JSON response to that file; parent directories are created."),
                  new CliSurface.DefinitionEntry(
                      "source.path", "open an existing workbook from that path."),
                  new CliSurface.DefinitionEntry(
                      "persistence SAVE_AS.path",
                      "write a new workbook to that path; parent directories are created."),
                  new CliSurface.DefinitionEntry(
                      "persistence OVERWRITE",
                      "write back to source.path; no path field is supplied."),
                  new CliSurface.DefinitionEntry(
                      "Relative CLI flag paths",
                      GridGrindContractText.cliFlagPathResolutionSummary()),
                  new CliSurface.DefinitionEntry(
                      "Relative request-owned paths",
                      "source.path, persistence paths, source-backed file inputs,"
                          + " formulaEnvironment.externalWorkbooks[*].path, and"
                          + " persistence.security.signature.pkcs12Path follow one rule:"
                          + " "
                          + GridGrindContractText.requestOwnedPathResolutionSummary()),
                  new CliSurface.DefinitionEntry(
                      "Relative FILE hyperlink targets",
                      "are analyzed against the persisted workbook path when one exists; use"
                          + " absolute paths for cwd-independent results."))),
          new CliSurface.CliTableSection(
              "Coordinate Systems",
              "Pattern",
              "Convention / Example",
              List.of(
                  new CliSurface.CoordinateSystemEntry("address", "A1 cell address, e.g. B3"),
                  new CliSurface.CoordinateSystemEntry("range", "A1 rectangular range, e.g. A1:C4"),
                  new CliSurface.CoordinateSystemEntry(
                      "*RowIndex", "zero-based, e.g. 0 = Excel row 1"),
                  new CliSurface.CoordinateSystemEntry(
                      "*ColumnIndex", "zero-based, e.g. 0 = Excel column A"),
                  new CliSurface.CoordinateSystemEntry(
                      "first/last pairs", "inclusive zero-based bands."))),
          new CliSurface.CliTemplateSection("Minimal Valid Request"),
          new CliSurface.CliCommandExample(
              "stdin Example", List.of("gridgrind --print-request-template | gridgrind"), null),
          new CliSurface.CliCommandExample(
              "Docker File Example",
              List.of(
                  "docker run --rm -i \\",
                  "  -v \"$(pwd)\":/workdir \\",
                  "  -w /workdir \\",
                  "  {{CONTAINER_TAG}} \\",
                  "  --request request.json \\",
                  "  --response response.json"),
              "In Docker, mount the host directory that contains your request and workbook"
                  + " files, then set -w to that mount point so every relative path resolves"
                  + " inside the mounted directory."),
          new CliSurface.CliDiscoverySection(
              "Discovery",
              List.of(
                  "gridgrind --print-request-template",
                  "gridgrind --doctor-request",
                  "gridgrind --print-task-catalog",
                  "gridgrind --print-task-plan <id>",
                  "gridgrind --print-goal-plan \"monthly sales dashboard with charts\"",
                  "gridgrind --print-protocol-catalog",
                  GridGrindContractText.workbookFindingsDiscoverySummary()),
              "Built-in generated examples",
              "Print one built-in example",
              "The task catalog publishes high-level office-work recipes composed from exact"
                  + " protocol capabilities. The protocol catalog remains the authoritative"
                  + " execution contract: it lists each field, whether it is required, and"
                  + " the nested/plain type group accepted by polymorphic fields such as"
                  + " target, action, query, value, style, and scope. Mutation, assertion,"
                  + " and inspection entries also publish targetSelectors and/or"
                  + " targetSelectorRule so agents can see the allowed target families,"
                  + " derived selector rules, and any shared-selector disambiguation notes"
                  + " before sending a request. When ids repeat across groups, qualify the"
                  + " lookup as <group>:<id> such as cellInputTypes:FORMULA. Nested and"
                  + " plain type-group descriptors can also be queried directly, for example"
                  + " nestedTypes:cellInputTypes or plainTypes:chartInputType.",
              "gridgrind --print-example " + GridGrindShippedExamples.examples().getFirst().id()),
          new CliSurface.CliReferenceSection(
              "Docs",
              List.of(
                  new CliSurface.ReferenceEntry("Quick reference", "docs/QUICK_REFERENCE.md"),
                  new CliSurface.ReferenceEntry("Operations reference", "docs/OPERATIONS.md"),
                  new CliSurface.ReferenceEntry("Error reference", "docs/ERRORS.md"))),
          new CliSurface.CliDefinitionSection(
              "Flags",
              List.of(
                  new CliSurface.DefinitionEntry(
                      "--request <path>", "Read the JSON request from a file instead of stdin."),
                  new CliSurface.DefinitionEntry(
                      "--response <path>", "Write the JSON response to a file instead of stdout."),
                  new CliSurface.DefinitionEntry(
                      "--doctor-request",
                      "Lint one request, preflight source-backed input resolution plus existing"
                          + " workbook-source accessibility, and emit a machine-readable doctor"
                          + " report without mutating a workbook."),
                  new CliSurface.DefinitionEntry(
                      "--print-request-template", "Print a minimal valid request JSON document."),
                  new CliSurface.DefinitionEntry(
                      "--print-task-catalog",
                      "Print the machine-readable task catalog of high-level office-work"
                          + " recipes."),
                  new CliSurface.DefinitionEntry(
                      "--task <id>",
                      "With --print-task-catalog, print one task entry by its stable id."),
                  new CliSurface.DefinitionEntry(
                      "--print-task-plan <id>",
                      "Print a machine-readable starter request scaffold for one task id."),
                  new CliSurface.DefinitionEntry(
                      "--print-goal-plan <goal>",
                      "Print ranked contract-owned task matches plus starter scaffolds for one"
                          + " freeform goal."),
                  new CliSurface.DefinitionEntry(
                      "--print-protocol-catalog", "Print the machine-readable protocol catalog."),
                  new CliSurface.DefinitionEntry(
                      "--operation <id>",
                      "With --print-protocol-catalog, print one unique entry or one nested/"
                          + " plain type group; qualify duplicate ids as <group>:<id> and"
                          + " query type groups as nestedTypes:<group> or plainTypes:<group>."),
                  new CliSurface.DefinitionEntry(
                      "--search <text>",
                      "With --print-protocol-catalog, perform case-insensitive search across"
                          + " lookup ids, qualified ids, catalog groups, and summaries."),
                  new CliSurface.DefinitionEntry(
                      "--print-example <id>", "Print one built-in generated example request."),
                  new CliSurface.DefinitionEntry("--help, -h, help", "Print this help text."),
                  new CliSurface.DefinitionEntry(
                      "--version", "Print the GridGrind version and description."),
                  new CliSurface.DefinitionEntry(
                      "--license", "Print the GridGrind license and third-party notices."))),
          GridGrindContractText.standardInputRequiresRequestMessage());

  private GridGrindProtocolCatalogCliSurface() {}
}
