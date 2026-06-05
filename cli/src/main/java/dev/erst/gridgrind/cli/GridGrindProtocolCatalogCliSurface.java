package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.cli.examples.GridGrindShippedExamples;
import dev.erst.gridgrind.contract.catalog.GridGrindContractText;
import dev.erst.gridgrind.contract.catalog.GridGrindExecutionModeMetadata;
import java.util.List;
import java.util.Optional;

/** Owns the CLI/help surface published by the command-line transport. */
final class GridGrindProtocolCatalogCliSurface {
  static final CliSurface CLI_SURFACE =
      new CliSurface(
          new CliSurface.CliSection(
              "Usage",
              List.of(
                  "gridgrind --request <path> [--temp-root <path>] [--response <path>]",
                  "gridgrind --execution-root <path> [--temp-root <path>] [--response <path>]"
                      + " < request.json",
                  "gridgrind --doctor-request --request <path> [--temp-root <path>]"
                      + " [--response <path>]",
                  "gridgrind --doctor-request --execution-root <path> [--temp-root <path>]"
                      + " [--response <path>] < request.json",
                  "gridgrind --print-request-template [--response <path>]",
                  "gridgrind --print-example-catalog [--response <path>]",
                  "gridgrind --print-task-catalog [--lookup <id>] [--response <path>]",
                  "gridgrind --print-task-plan --lookup <id> [--response <path>]",
                  "gridgrind --print-task-keyword-match --query <text> [--response <path>]",
                  "gridgrind --print-protocol-catalog [--response <path>]",
                  "gridgrind --print-protocol-catalog --lookup <id>|<group>:<id> [--response"
                      + " <path>]",
                  "gridgrind --print-protocol-catalog --search <text> [--response <path>]",
                  "gridgrind --print-example --lookup <id> [--response <path>]",
                  "gridgrind --help | -h [--response <path>]",
                  "gridgrind --help-protocol [--response <path>]",
                  "gridgrind --help-guidance [--response <path>]",
                  "gridgrind --version [--response <path>]",
                  "gridgrind --license [--response <path>]")),
          new CliSurface.CliWorkflowSection(
              "Workflows",
              List.of(
                  new CliSurface.WorkflowEntry(
                      "Discover What To Send",
                      List.of(
                          "List shipped task recipes: gridgrind --print-task-catalog"
                              + " --response tasks.json",
                          "Get one executable starter scenario: gridgrind --print-task-plan"
                              + " --lookup DASHBOARD --response dashboard-request.json",
                          "Search exact protocol shapes: gridgrind --print-protocol-catalog"
                              + " --search \"chart title\" --response catalog-search.json",
                          "Resolve one exact catalog group: gridgrind --print-protocol-catalog"
                              + " --lookup calculationStrategyTypes --response"
                              + " calculation-strategy-types.json")),
                  new CliSurface.WorkflowEntry(
                      "Draft And Preflight",
                      List.of(
                          "Start from the minimal request: gridgrind --print-request-template"
                              + " --response request.json",
                          "For stdin-driven execution or doctoring, pass one explicit"
                              + " --execution-root so request-owned paths and temp files"
                              + " stay invocation-owned instead of ambient.",
                          "Use --print-task-plan --lookup <id> when you want one executable"
                              + " starter scenario instead of building from scratch.",
                          "Or copy one built-in example: gridgrind --print-example"
                              + " --lookup WORKBOOK_HEALTH --response example.json",
                          "Lint before executing: gridgrind --doctor-request --request"
                              + " request.json --response doctor.json")),
                  new CliSurface.WorkflowEntry(
                      "Execute And Keep Artifacts",
                      List.of(
                          "Run locally and capture the response: gridgrind --request"
                              + " request.json --response response.json",
                          "Reopen a saved workbook by switching source.type to EXISTING and"
                              + " pointing source.path at the .xlsx file you want to inspect"
                              + " or mutate.",
                          "Use persistence.type=SAVE_AS when the workbook should be written,"
                              + " and persistence.type=NONE when the run is read-only or"
                              + " diagnostic.")))),
          new CliSurface.CliSection(
              "Execution",
              List.of(
                  "GridGrind executes ordered steps in sequence, then saves the workbook (unless"
                      + " persistence is NONE); if any step fails, no workbook is written.",
                  "source.type=NEW starts with zero sheets; ENSURE_SHEET creates the first"
                      + " sheet.",
                  "execution is explicit; the request supplies execution.mode,"
                      + " execution.journal, and execution.calculation.")),
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
                      "Request JSON size",
                      "request JSON must not exceed 16 MiB ("
                          + GridGrindContractText.requestDocumentLimitBytes()
                          + " bytes); authored text and binary payloads may instead use"
                          + " UTF8_FILE, FILE, or STANDARD_INPUT sources."),
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
                      "rejected when they would move tables, sheet autofilters, or data"
                          + " validations; deletes/shifts also reject destructive range-backed"
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
                  "protocolVersion is required.",
                  "source.type is required. NEW creates a blank workbook."
                      + " EXISTING requires source.path.",
                  "persistence is required. NONE keeps the workbook in memory only.",
                  "planId is optional. When omitted, the response journal generates a"
                      + " plan-<uuid> correlation id.",
                  "execution is required. execution.mode is a typed discriminator;"
                      + " choose type=FULL_XSSF, EVENT_READ, or STREAMING_WRITE under the"
                      + " limits above.",
                  "execution.journal.level controls journal detail; VERBOSE also streams live"
                      + " phase events to stderr as timestamped CATEGORY detail lines with"
                      + " optional stepIndex/stepId pairs.",
                  "execution.calculation controls server-side evaluation, cache clearing, and"
                      + " open-time recalc flags. strategy.type accepts DO_NOT_CALCULATE,"
                      + " EVALUATE_ALL, EVALUATE_TARGETS, and CLEAR_CACHES_ONLY."
                      + " EVALUATE_TARGETS also requires strategy.cells[]."
                      + " markRecalculateOnOpen is independent.",
                  "Response telemetry is split intentionally: journal.* captures execution-phase"
                      + " timing and event chronology, while root calculation/warnings/assertions/"
                      + "inspections carry the authoritative outcome payloads.",
                  "journal.events[*].timestamp plus every durationMillis field are observational"
                      + " runtime telemetry and vary between runs; assert status/category/detail,"
                      + " not literal timings.",
                  "Source-backed text and binary fields support INLINE, UTF8_FILE or FILE, and"
                      + " STANDARD_INPUT sources.",
                  GridGrindContractText.standardInputRequiresRequestMessage()
                      + " because stdin cannot carry both the request JSON and authored input"
                      + " content in one CLI invocation.",
                  "formulaEnvironment is required. The empty-environment shape is"
                      + " externalWorkbooks=[], missingWorkbookPolicy=ERROR, udfToolpacks=[]."
                      + " missingWorkbookPolicy accepts ERROR or USE_CACHED_VALUE;"
                      + " udfToolpacks[] registers named UDF packs for formula evaluation.",
                  "array-formula braces such as {=SUM(A1:A2*B1:B2)} are rejected as"
                      + " INVALID_FORMULA.",
                  "steps is required. [] represents an empty no-op plan.",
                  GridGrindContractText.stepKindSummary()
                      + " target is a sibling field on each step; it is not nested inside"
                      + " action, assertion, or query.",
                  "Step order is authoritative: mutations, assertions, and inspections may be"
                      + " interleaved when the workflow demands it.")),
          new CliSurface.CliDefinitionSection(
              "File Workflow",
              List.of(
                  new CliSurface.DefinitionEntry(
                      "No --request flag",
                      "read the JSON request from stdin; pass --execution-root so relative"
                          + " request-owned paths and execution temp files resolve from one"
                          + " explicit directory. A bare TTY invocation with no piped request"
                          + " is rejected."),
                  new CliSurface.DefinitionEntry(
                      "--request <path>",
                      "read the JSON request from that file; the request file directory owns"
                          + " request-root resolution."),
                  new CliSurface.DefinitionEntry(
                      "--execution-root <path>",
                      "required when the request JSON arrives on stdin; relative request-owned"
                          + " paths resolve from that directory."),
                  new CliSurface.DefinitionEntry(
                      "--temp-root <path>",
                      "override execution scratch space. Without it, temp files resolve under"
                          + " .gridgrind/tmp inside the request root or explicit"
                          + " --execution-root."),
                  new CliSurface.DefinitionEntry(
                      "No --response flag", "write the primary command output to stdout."),
                  new CliSurface.DefinitionEntry(
                      "--response <path>",
                      "write the primary command output to that file; parent directories are"
                          + " created. Execution writes the JSON response, doctoring writes"
                          + " the doctor report, and help or discovery commands write their"
                          + " rendered text or JSON payload. Non-success results also emit one"
                          + " stderr pointer line naming the file: CLI argument errors write"
                          + " 'CLI failure report', request-content errors write 'request"
                          + " failure report', execution failures write 'response', and doctor"
                          + " failures write 'doctor report'."),
                  new CliSurface.DefinitionEntry(
                      "source.type=EXISTING + source.path",
                      "open an existing workbook from that path."),
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
              "Stdin Example",
              List.of("gridgrind --print-request-template | gridgrind --execution-root ."),
              Optional.empty()),
          new CliSurface.CliCommandExample(
              "Docker Example",
              List.of(
                  "docker run --rm -i \\",
                  "  -v \"$(pwd)\":/workdir \\",
                  "  -w /workdir \\",
                  "  {{CONTAINER_TAG}} \\",
                  "  --request request.json \\",
                  "  --response response.json"),
              Optional.of(
                  "In Docker, mount the host directory that contains your request and workbook"
                      + " files, then set -w to that mount point so every relative path resolves"
                      + " inside the mounted directory. From a repository checkout, build the"
                      + " same runtime surface with 'docker buildx build --load -t"
                      + " gridgrind-local .' and replace {{CONTAINER_TAG}} with"
                      + " 'gridgrind-local'.")),
          new CliSurface.CliDiscoverySection(
              "Discovery",
              List.of(
                  "gridgrind --print-request-template --response request.json",
                  "gridgrind --doctor-request --request request.json --response doctor.json",
                  "gridgrind --print-example-catalog --response example-catalog.json",
                  "gridgrind --print-task-catalog --response tasks.json",
                  "gridgrind --print-task-plan --lookup <id> --response task-plan.json",
                  "gridgrind --print-task-keyword-match --query \"monthly sales dashboard with charts\""
                      + " --response task-keyword-match.json",
                  "gridgrind --print-protocol-catalog --response protocol-catalog.json"),
              "Built-in generated examples",
              "Print one built-in example",
              List.of(
                  new CliSurface.WorkflowEntry(
                      "Example portability",
                      List.of(
                          "SELF_CONTAINED starters execute from a blank working directory.",
                          "REQUIRES_EXAMPLE_ASSETS starters require copied asset paths beside the"
                              + " request file; requiredPaths names those paths directly.",
                          GridGrindContractText.workbookFindingsDiscoverySummary()
                              + " Include it in any diagnostic plan with persistence.type=NONE.")),
                  new CliSurface.WorkflowEntry(
                      "Task starters",
                      List.of(
                          "The CLI task catalog publishes high-level office-work recipes composed"
                              + " from exact protocol capabilities.",
                          "Each task entry now publishes starter.suggestedRequestPath,"
                              + " starter.workspaceMode, and starter.requiredPaths so agents can"
                              + " decide whether one task starter is self-contained before"
                              + " printing it.")),
                  new CliSurface.WorkflowEntry(
                      "Protocol catalog search",
                      List.of(
                          "The protocol catalog remains the authoritative execution contract: it"
                              + " lists each field, whether it is required, and the nested/plain"
                              + " type group accepted by polymorphic fields such as target,"
                              + " action, query, value, style, and scope.",
                          "Search output is summary-first: it lists ids, summaries, related entry"
                              + " ids, and supporting ids. Rerun --lookup when you need one full"
                              + " entry or type-group definition.",
                          "Unqualified --lookup resolves individual type ids (SET_CELL,"
                              + " ENSURE_SHEET, GET_CELLS, EXPECT_CELL_VALUE, …), nested/plain"
                              + " type-group names (cellInputTypes, calculationStrategyTypes, …),"
                              + " and top-level operation category names (mutationActionTypes,"
                              + " assertionTypes, inspectionQueryTypes, sourceTypes,"
                              + " persistenceTypes, stepTypes). Qualify with nestedTypes: or"
                              + " plainTypes: only when ids repeat across groups."))),
              "gridgrind --print-example --lookup "
                  + GridGrindShippedExamples.catalog().examples().getFirst().id()
                  + " --response example.json"),
          new CliSurface.CliReferenceSection(
              "Docs",
              List.of(
                  new CliSurface.ReferenceEntry(
                      "Quick reference",
                      "docs/QUICK_REFERENCE.md",
                      "Synopsis-first command cheat sheet for first-contact lookup."),
                  new CliSurface.ReferenceEntry(
                      "Operations reference",
                      "docs/OPERATIONS.md",
                      "Longer operator workflows, artifact handling, and runtime/container use."),
                  new CliSurface.ReferenceEntry(
                      "Error reference",
                      "docs/ERRORS.md",
                      "Problem codes, failure interpretation, and recovery guidance."))),
          new CliSurface.CliDefinitionSection(
              "Flags",
              List.of(
                  new CliSurface.DefinitionEntry(
                      "--request <path>", "Read the JSON request from a file instead of stdin."),
                  new CliSurface.DefinitionEntry(
                      "--execution-root <path>",
                      "Required when the request JSON arrives on stdin; relative request-owned"
                          + " paths and execution temp files resolve from that directory."),
                  new CliSurface.DefinitionEntry(
                      "--temp-root <path>",
                      "Override execution scratch space. Without it, GridGrind uses"
                          + " .gridgrind/tmp inside the request root or explicit"
                          + " --execution-root."),
                  new CliSurface.DefinitionEntry(
                      "--response <path>",
                      "Write the primary command output to a file instead of stdout."),
                  new CliSurface.DefinitionEntry(
                      "--doctor-request",
                      "Lint one request, preflight source-backed input resolution plus existing"
                          + " workbook-source accessibility, and emit a machine-readable doctor"
                          + " report without mutating a workbook. The doctor response returns"
                          + " warnings plus every independently provable blocking problem."),
                  new CliSurface.DefinitionEntry(
                      "--print-request-template", "Print a minimal valid request JSON document."),
                  new CliSurface.DefinitionEntry(
                      "--print-example-catalog",
                      "Print the machine-readable built-in example catalog, including the stable"
                          + " workspaceMode portability contract plus any requiredPaths for"
                          + " asset-backed example ids."),
                  new CliSurface.DefinitionEntry(
                      "--print-task-catalog",
                      "Print the machine-readable task catalog of high-level office-work"
                          + " recipes, including starter.suggestedRequestPath,"
                          + " starter.workspaceMode, and starter.requiredPaths."),
                  new CliSurface.DefinitionEntry(
                      "--lookup <id>",
                      "With --print-example, --print-task-catalog, --print-task-plan, or"
                          + " --print-protocol-catalog, print one stable entry by id (SET_CELL,"
                          + " ENSURE_SHEET, …), one nested/plain type-group descriptor by group"
                          + " name (cellInputTypes, calculationStrategyTypes, …), or one"
                          + " top-level category array by name (mutationActionTypes,"
                          + " assertionTypes, inspectionQueryTypes, sourceTypes,"
                          + " persistenceTypes, stepTypes). Qualify as <category>:<id>"
                          + " (e.g. mutationActionTypes:SET_CELL) when ids repeat across groups."),
                  new CliSurface.DefinitionEntry(
                      "--print-task-plan --lookup <id>",
                      "Print one executable starter scenario for one task id."),
                  new CliSurface.DefinitionEntry(
                      "--print-task-keyword-match --query <text>",
                      "Print ranked CLI-owned task matches for one English keyword query."
                          + " Use --print-task-plan --lookup <id> for the executable starter"
                          + " scenario after you choose a task id. After normalization, at least"
                          + " one searchable non-stop-word term must remain."),
                  new CliSurface.DefinitionEntry(
                      "--print-protocol-catalog", "Print the machine-readable protocol catalog."),
                  new CliSurface.DefinitionEntry(
                      "--search <text>",
                      "With --print-protocol-catalog, perform case-insensitive search across"
                          + " lookup ids, qualified ids, catalog groups, and summaries."
                          + " Search promotes top-level operations first and uses"
                          + " relatedEntryIds on support-group hits."),
                  new CliSurface.DefinitionEntry(
                      "--print-example --lookup <id>",
                      "Print one built-in generated example request."),
                  new CliSurface.DefinitionEntry("--help, -h", "Print the short synopsis."),
                  new CliSurface.DefinitionEntry(
                      "--help-protocol", "Print the authoritative CLI and request grammar only."),
                  new CliSurface.DefinitionEntry(
                      "--help-guidance",
                      "Print workflow guidance, examples, Docker usage, and discovery playbooks."),
                  new CliSurface.DefinitionEntry(
                      "--version", "Print the GridGrind version and description."),
                  new CliSurface.DefinitionEntry(
                      "--license", "Print the GridGrind license and third-party notices."))),
          GridGrindContractText.standardInputRequiresRequestMessage());

  private GridGrindProtocolCatalogCliSurface() {}
}
