package dev.erst.gridgrind.contract.catalog;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.catalog.gather.CatalogDuplicateFailures;
import dev.erst.gridgrind.contract.catalog.gather.CatalogFieldMetadataSupport;
import dev.erst.gridgrind.contract.catalog.gather.CatalogGatherers;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.contract.step.WorkbookStep;
import dev.erst.gridgrind.contract.step.WorkbookStepTargeting;
import java.lang.reflect.RecordComponent;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.Set;
import java.util.function.Function;

/**
 * Publishes the machine-readable GridGrind protocol surface used by CLI discovery commands.
 *
 * <p>This registry intentionally spans the full public contract surface, so PMD's import-count
 * heuristic is not a useful coupling signal here.
 */
public final class GridGrindProtocolCatalog {
  private static final String DISCRIMINATOR_FIELD = "type";
  private static final WorkbookPlan REQUEST_TEMPLATE =
      new WorkbookPlan(
          GridGrindProtocolVersion.current(),
          null,
          new WorkbookPlan.WorkbookSource.New(),
          new WorkbookPlan.WorkbookPersistence.None(),
          null,
          null,
          List.of());
  private static final TypeEntry REQUEST_TYPE =
      typeEntry(
          WorkbookPlan.class,
          "WorkbookPlan",
          "Complete GridGrind plan for workbook source, execution policy, formula environment,"
              + " ordered mutation, assertion, and inspection steps, and persistence.",
          List.of(
              "protocolVersion",
              "planId",
              "persistence",
              "execution",
              "formulaEnvironment",
              "steps"));
  private static final CliSurface CLI_SURFACE =
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
                  "gridgrind --help | -h",
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
                      "Relative paths",
                      "in --request, --response, source.path, and persistence.path resolve from"
                          + " the current working directory."),
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
                      "Lint one request and emit a machine-readable doctor report without"
                          + " opening or mutating a workbook."),
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
                  new CliSurface.DefinitionEntry("--help, -h", "Print this help text."),
                  new CliSurface.DefinitionEntry(
                      "--version", "Print the GridGrind version and description."),
                  new CliSurface.DefinitionEntry(
                      "--license", "Print the GridGrind license and third-party notices."))),
          GridGrindContractText.standardInputRequiresRequestMessage());
  private static final List<CatalogTypeDescriptor> STEP_TYPES =
      List.of(
          descriptor(
              MutationStep.class,
              "MUTATION",
              "Execute one mutation action against the selected workbook target."),
          descriptor(
              AssertionStep.class,
              "ASSERTION",
              "Verify one authored expectation against the selected workbook target."),
          descriptor(
              InspectionStep.class,
              "INSPECTION",
              "Run one factual or analytical inspection query against the selected workbook"
                  + " target."));
  private static final List<CatalogTypeDescriptor> SOURCE_TYPES =
      List.of(
          descriptor(
              WorkbookPlan.WorkbookSource.New.class,
              "NEW",
              "Create a brand-new empty workbook. A new workbook starts with zero sheets;"
                  + " use ENSURE_SHEET to create the first sheet."),
          descriptor(
              WorkbookPlan.WorkbookSource.ExistingFile.class,
              "EXISTING",
              "Open an existing .xlsx workbook from disk."
                  + " Relative paths resolve in the current execution environment."
                  + " source.security.password unlocks encrypted OOXML packages.",
              "security"));
  private static final List<CatalogTypeDescriptor> PERSISTENCE_TYPES =
      List.of(
          descriptor(
              WorkbookPlan.WorkbookPersistence.None.class,
              "NONE",
              "Keep the workbook in memory only." + " The response persistence.type echoes NONE."),
          descriptor(
              WorkbookPlan.WorkbookPersistence.OverwriteSource.class,
              "OVERWRITE",
              "Overwrite the opened source workbook at source.path."
                  + " No path field is accepted on OVERWRITE;"
                  + " the write target is the same path opened by the EXISTING source."
                  + " Relative source.path values resolve in the current execution environment."
                  + " persistence.security can encrypt and/or sign the saved OOXML package."
                  + " The response persistence.type echoes OVERWRITE and includes sourcePath"
                  + " (the original source path string) and executionPath (absolute normalized).",
              "security"),
          descriptor(
              WorkbookPlan.WorkbookPersistence.SaveAs.class,
              "SAVE_AS",
              "Save the workbook to a new .xlsx path."
                  + " Relative paths resolve in the current execution environment."
                  + " persistence.security can encrypt and/or sign the saved OOXML package."
                  + " The response persistence.type echoes SAVE_AS and includes requestedPath"
                  + " (the literal path from the request) and executionPath (the absolute"
                  + " normalized path where the file was written); they differ when a relative"
                  + " path or a path with .. segments is supplied."
                  + " Missing parent directories are created automatically.",
              "security"));
  private static final List<CatalogTypeDescriptor> MUTATION_ACTION_TYPES =
      List.of(
          descriptor(
              MutationAction.EnsureSheet.class,
              "ENSURE_SHEET",
              "Create the sheet if it does not already exist."
                  + " Sheet names must be 1 to 31 characters and must not contain"
                  + " : \\ / ? * [ ] or begin or end with a single quote (Excel limit)."),
          descriptor(
              MutationAction.RenameSheet.class,
              "RENAME_SHEET",
              "Rename an existing sheet."
                  + " The new name must be 1 to 31 characters and must not contain"
                  + " : \\ / ? * [ ] or begin or end with a single quote (Excel limit)."),
          descriptor(
              MutationAction.DeleteSheet.class,
              "DELETE_SHEET",
              "Delete an existing sheet."
                  + " A workbook must retain at least one sheet and at least one visible sheet;"
                  + " deleting the last sheet or the last visible sheet returns INVALID_REQUEST."),
          descriptor(
              MutationAction.MoveSheet.class,
              "MOVE_SHEET",
              "Move a sheet to a zero-based workbook position."
                  + " targetIndex is 0-based: 0 moves the sheet to the front,"
                  + " sheetCount-1 moves it to the back."),
          descriptor(
              MutationAction.CopySheet.class,
              "COPY_SHEET",
              "Copy one sheet into a new visible, unselected sheet."
                  + " newSheetName follows the same Excel sheet-name rules as ENSURE_SHEET."
                  + " position defaults to APPEND_AT_END when omitted."
                  + " The copied sheet preserves sheet-local workbook content such as formulas,"
                  + " validations, conditional formatting, comments, hyperlinks, merged regions,"
                  + " layout state, sheet-scoped named ranges, sheet autofilters, and tables."
                  + " Copied tables are renamed automatically when needed to keep workbook-global"
                  + " table names unique.",
              "position"),
          descriptor(
              MutationAction.SetActiveSheet.class,
              "SET_ACTIVE_SHEET",
              "Set the active sheet."
                  + " Hidden sheets cannot be activated."
                  + " The active sheet is always selected."),
          descriptor(
              MutationAction.SetSelectedSheets.class,
              "SET_SELECTED_SHEETS",
              "Set the selected visible sheet set."
                  + " Duplicate or unknown sheet names are rejected."
                  + " selectedSheetNames read back in workbook order;"
                  + " activeSheetName preserves the primary selected sheet choice."),
          descriptor(
              MutationAction.SetSheetVisibility.class,
              "SET_SHEET_VISIBILITY",
              "Set one sheet visibility state."
                  + " A workbook must retain at least one visible sheet;"
                  + " hiding the last visible sheet is rejected."),
          descriptor(
              MutationAction.SetSheetProtection.class,
              "SET_SHEET_PROTECTION",
              "Enable sheet protection with the exact supported lock flags."
                  + " password is optional; when provided it is hashed into the sheet-protection"
                  + " metadata.",
              "password"),
          descriptor(
              MutationAction.ClearSheetProtection.class,
              "CLEAR_SHEET_PROTECTION",
              "Disable sheet protection entirely."),
          descriptor(
              MutationAction.SetWorkbookProtection.class,
              "SET_WORKBOOK_PROTECTION",
              "Enable workbook-level protection and optional workbook or revisions passwords."
                  + " Omitted lock flags normalize to false and omitted passwords are cleared."),
          descriptor(
              MutationAction.ClearWorkbookProtection.class,
              "CLEAR_WORKBOOK_PROTECTION",
              "Clear workbook-level protection and any stored workbook or revisions passwords."),
          descriptor(
              MutationAction.MergeCells.class,
              "MERGE_CELLS",
              "Merge a rectangular A1-style range."),
          descriptor(
              MutationAction.UnmergeCells.class,
              "UNMERGE_CELLS",
              "Remove one merged region by exact range match."),
          descriptor(
              MutationAction.SetColumnWidth.class,
              "SET_COLUMN_WIDTH",
              "Set one or more column widths in Excel character units."
                  + " widthCharacters must be > 0 and <= 255 (Excel column width limit)."),
          descriptor(
              MutationAction.SetRowHeight.class,
              "SET_ROW_HEIGHT",
              "Set one or more row heights in Excel point units."
                  + " heightPoints must be > 0 and <= "
                  + dev.erst.gridgrind.excel.foundation.ExcelSheetLayoutLimits.MAX_ROW_HEIGHT_POINTS
                  + " (Excel row height limit)."),
          descriptor(
              MutationAction.InsertRows.class,
              "INSERT_ROWS",
              "Insert one or more blank rows before rowIndex."
                  + " rowIndex must be <= last existing row + 1."
                  + " Append-edge inserts on sparse sheets do not materialize a new physical tail"
                  + " row until content or row metadata exists there."
                  + " GridGrind rejects row inserts that would move tables, sheet autofilters,"
                  + " or data validations."),
          descriptor(
              MutationAction.DeleteRows.class,
              "DELETE_ROWS",
              "Delete one inclusive zero-based row band."
                  + " GridGrind rejects row deletes that would move or truncate tables,"
                  + " sheet autofilters, or data validations, and it also rejects deletes that"
                  + " would truncate range-backed named ranges."),
          descriptor(
              MutationAction.ShiftRows.class,
              "SHIFT_ROWS",
              "Move one inclusive zero-based row band by delta rows."
                  + " delta must not be 0."
                  + " GridGrind rejects row shifts that would move tables, sheet autofilters,"
                  + " or data validations, and it also rejects shifts that would partially move"
                  + " or overwrite range-backed named ranges."),
          descriptor(
              MutationAction.InsertColumns.class,
              "INSERT_COLUMNS",
              "Insert one or more blank columns before columnIndex."
                  + " columnIndex must be <= last existing column + 1."
                  + " Append-edge inserts on sparse sheets do not materialize a new physical"
                  + " tail column until cells or explicit column metadata exist there."
                  + " GridGrind rejects column inserts when the workbook contains formula cells"
                  + " or formula-defined names, or when the edit would move tables,"
                  + " sheet autofilters, or data validations."),
          descriptor(
              MutationAction.DeleteColumns.class,
              "DELETE_COLUMNS",
              "Delete one inclusive zero-based column band."
                  + " GridGrind rejects column deletes when the workbook contains formula cells"
                  + " or formula-defined names, or when the edit would move or truncate tables,"
                  + " sheet autofilters, or data validations, or truncate range-backed named"
                  + " ranges."),
          descriptor(
              MutationAction.ShiftColumns.class,
              "SHIFT_COLUMNS",
              "Move one inclusive zero-based column band by delta columns."
                  + " delta must not be 0."
                  + " GridGrind rejects column shifts when the workbook contains formula cells"
                  + " or formula-defined names, or when the edit would move tables,"
                  + " sheet autofilters, or data validations, or partially move or overwrite"
                  + " range-backed named ranges."),
          descriptor(
              MutationAction.SetRowVisibility.class,
              "SET_ROW_VISIBILITY",
              "Set the hidden state for one inclusive zero-based row band."),
          descriptor(
              MutationAction.SetColumnVisibility.class,
              "SET_COLUMN_VISIBILITY",
              "Set the hidden state for one inclusive zero-based column band."),
          descriptor(
              MutationAction.GroupRows.class,
              "GROUP_ROWS",
              "Apply one outline group to an inclusive zero-based row band."
                  + " collapsed defaults to false when omitted."),
          descriptor(
              MutationAction.UngroupRows.class,
              "UNGROUP_ROWS",
              "Remove outline grouping from one inclusive zero-based row band."),
          descriptor(
              MutationAction.GroupColumns.class,
              "GROUP_COLUMNS",
              "Apply one outline group to an inclusive zero-based column band."
                  + " collapsed defaults to false when omitted."),
          descriptor(
              MutationAction.UngroupColumns.class,
              "UNGROUP_COLUMNS",
              "Remove outline grouping from one inclusive zero-based column band."),
          descriptor(
              MutationAction.SetSheetPane.class,
              "SET_SHEET_PANE",
              "Apply one explicit pane state."
                  + " pane.type can be NONE, FROZEN, or SPLIT; use NONE to clear panes."),
          descriptor(
              MutationAction.SetSheetZoom.class,
              "SET_SHEET_ZOOM",
              "Set the sheet zoom percentage."
                  + " zoomPercent must be between "
                  + dev.erst.gridgrind.excel.foundation.ExcelSheetLayoutLimits.MIN_ZOOM_PERCENT
                  + " and "
                  + dev.erst.gridgrind.excel.foundation.ExcelSheetLayoutLimits.MAX_ZOOM_PERCENT
                  + " inclusive."),
          descriptor(
              MutationAction.SetSheetPresentation.class,
              "SET_SHEET_PRESENTATION",
              "Apply one authoritative supported sheet-presentation state to a sheet."
                  + " Omitted nested fields normalize to defaults or clear state."
                  + " The supported surface covers screen display flags, right-to-left mode,"
                  + " tab color, outline summary placement, authored default row and column sizing"
                  + " (defaultColumnWidth > 0 and <= 255; defaultRowHeightPoints > 0 and <= "
                  + dev.erst.gridgrind.excel.foundation.ExcelSheetLayoutLimits.MAX_ROW_HEIGHT_POINTS
                  + "), and ignored-errors ranges."),
          descriptor(
              MutationAction.SetPrintLayout.class,
              "SET_PRINT_LAYOUT",
              "Apply one authoritative supported print-layout state to a sheet."
                  + " Omitted nested fields normalize to default or clear state."
                  + " The supported surface covers print area, orientation, fit scaling,"
                  + " repeating rows, repeating columns, plain header or footer text,"
                  + " margins, printGridlines, centering, paper size, draft,"
                  + " black-and-white, copies,"
                  + " first-page numbering, and explicit row or column breaks."),
          descriptor(
              MutationAction.ClearPrintLayout.class,
              "CLEAR_PRINT_LAYOUT",
              "Clear the supported print-layout state from a sheet."),
          descriptor(
              MutationAction.SetCell.class, "SET_CELL", "Write one typed value to a single cell."),
          descriptor(
              MutationAction.SetRange.class,
              "SET_RANGE",
              "Write a rectangular grid of typed values."),
          descriptor(
              MutationAction.SetArrayFormula.class,
              "SET_ARRAY_FORMULA",
              "Author one contiguous single-cell or multi-cell array-formula group."
                  + " source accepts inline formula text with or without leading = or {=...}"
                  + " wrapper syntax."),
          descriptor(
              MutationAction.ClearArrayFormula.class,
              "CLEAR_ARRAY_FORMULA",
              "Remove the stored array-formula group targeted by any member cell."
                  + " Non-array cells are rejected explicitly."),
          descriptor(
              MutationAction.ImportCustomXmlMapping.class,
              "IMPORT_CUSTOM_XML_MAPPING",
              "Import XML content into one existing workbook custom-XML mapping."
                  + " mapping locates one existing map by mapId and/or name and xml accepts"
                  + " inline, file-backed, or STANDARD_INPUT text sources."
                  + " Workbook custom-XML mappings themselves must already exist in the source"
                  + " workbook; GridGrind imports data into them but does not author new map"
                  + " definitions.",
              "mapping"),
          descriptor(
              MutationAction.ClearRange.class,
              "CLEAR_RANGE",
              "Clear value, style, hyperlink, and comment state from a range."),
          descriptor(
              MutationAction.SetHyperlink.class,
              "SET_HYPERLINK",
              "Attach a hyperlink to one cell."
                  + " FILE targets use the field name path and normalize file: URIs to plain"
                  + " path strings."
                  + " Relative FILE targets are analyzed against the persisted workbook path"
                  + " when one exists."),
          descriptor(
              MutationAction.ClearHyperlink.class,
              "CLEAR_HYPERLINK",
              "Remove the hyperlink from one cell; no-op when the cell does not physically exist."),
          descriptor(
              MutationAction.SetComment.class,
              "SET_COMMENT",
              "Attach a comment to one cell."
                  + " Comments can carry ordered rich-text runs and an explicit anchor box;"
                  + " runs must concatenate to text."),
          descriptor(
              MutationAction.ClearComment.class,
              "CLEAR_COMMENT",
              "Remove the comment from one cell; no-op when the cell does not physically exist."),
          descriptor(
              MutationAction.SetPicture.class,
              "SET_PICTURE",
              "Create or replace one named picture on a sheet."
                  + " anchor uses the explicit DrawingAnchorInput discriminated shape and"
                  + " currently supports TWO_CELL anchors."
                  + " Reusing an existing object name replaces that picture authoritatively."),
          descriptor(
              MutationAction.SetSignatureLine.class,
              "SET_SIGNATURE_LINE",
              "Create or replace one named signature line on a sheet."
                  + " Signature lines surface through GET_DRAWING_OBJECTS and reuse"
                  + " SET_DRAWING_OBJECT_ANCHOR / DELETE_DRAWING_OBJECT for later edits."
                  + " anchor uses the explicit DrawingAnchorInput discriminated shape and"
                  + " currently supports TWO_CELL anchors."
                  + " Reusing an existing object name replaces any prior drawing object of that"
                  + " name authoritatively."),
          descriptor(
              MutationAction.SetChart.class,
              "SET_CHART",
              "Create or mutate one named chart on a sheet."
                  + " Supported authored families are AREA, AREA_3D, BAR, BAR_3D, DOUGHNUT,"
                  + " LINE, LINE_3D, PIE, PIE_3D, RADAR, SCATTER, SURFACE, and SURFACE_3D."
                  + " series bind to contiguous ranges or defined names and anchors currently"
                  + " support only TWO_CELL."
                  + " Chart and series FORMULA titles must resolve to one cell."
                  + " Failed validation leaves existing drawing state unchanged and creates no"
                  + " partial chart artifacts."
                  + " Existing unsupported chart detail is preserved on unrelated edits and"
                  + " rejected for authoritative mutation."),
          descriptor(
              MutationAction.SetPivotTable.class,
              "SET_PIVOT_TABLE",
              "Create or replace one workbook-global pivot table to POI XSSF's supported"
                  + " limited extent."
                  + " rowLabels, columnLabels, reportFilters, and dataFields must use disjoint"
                  + " source columns because POI persists only one role per pivot field."
                  + " When reportFilters are present, anchor.topLeftAddress must be on row 3 or"
                  + " lower so Excel's page-filter layout has room above the rendered body."),
          descriptor(
              MutationAction.SetShape.class,
              "SET_SHAPE",
              "Create or replace one named authored drawing shape on a sheet."
                  + " kind currently supports only SIMPLE_SHAPE and CONNECTOR."
                  + " SIMPLE_SHAPE defaults presetGeometryToken to rect when omitted."
                  + " anchor uses the explicit DrawingAnchorInput discriminated shape and"
                  + " currently supports TWO_CELL anchors."
                  + " Failed validation leaves existing drawing state unchanged and creates no"
                  + " partial shape artifacts."),
          descriptor(
              MutationAction.SetEmbeddedObject.class,
              "SET_EMBEDDED_OBJECT",
              "Create or replace one named embedded OLE package on a sheet."
                  + " previewImage supplies the visible preview raster required by Excel"
                  + " and Apache POI."
                  + " anchor uses the explicit DrawingAnchorInput discriminated shape and"
                  + " currently supports TWO_CELL anchors."),
          descriptor(
              MutationAction.SetDrawingObjectAnchor.class,
              "SET_DRAWING_OBJECT_ANCHOR",
              "Move one existing named picture, signature line, connector, simple shape,"
                  + " chart frame, or embedded object to a new authored anchor."
                  + " anchor currently supports only TWO_CELL anchors."
                  + " Read-only loaded families such as groups and graphic frames are rejected."),
          descriptor(
              MutationAction.DeleteDrawingObject.class,
              "DELETE_DRAWING_OBJECT",
              "Delete one existing named drawing object from the sheet."
                  + " Package relationships for picture media, signature-line preview images,"
                  + " and embedded-object parts are"
                  + " cleaned up when no other drawing object still references them."),
          descriptor(
              MutationAction.ApplyStyle.class,
              "APPLY_STYLE",
              "Apply a style patch to every cell in a range."
                  + " Write shape: style.border is a nested object"
                  + " { \"all\": { \"style\": \"THIN\" } } or per-side top/right/bottom/left."
                  + " Read shape (GET_CELLS, GET_WINDOW): borders are flat properties"
                  + " topBorderStyle, rightBorderStyle, bottomBorderStyle, leftBorderStyle;"
                  + " the nested border object is write-only."),
          descriptor(
              MutationAction.SetDataValidation.class,
              "SET_DATA_VALIDATION",
              "Create or replace one data-validation rule over the supplied sheet range."
                  + " Overlapping existing rules are normalized so the written rule becomes"
                  + " authoritative on its target range."),
          descriptor(
              MutationAction.ClearDataValidations.class,
              "CLEAR_DATA_VALIDATIONS",
              "Remove data-validation structures from the selected ranges on one sheet."
                  + " SELECTED removes only intersecting coverage; ALL clears every rule"
                  + " on the sheet."),
          descriptor(
              MutationAction.SetConditionalFormatting.class,
              "SET_CONDITIONAL_FORMATTING",
              "Create or replace one logical conditional-formatting block over the supplied"
                  + " sheet ranges."
                  + " The write contract authors formula rules, cell-value rules, color scales,"
                  + " data bars, icon sets, and top/bottom N rules."
                  + " Any existing conditional-formatting block that intersects the target ranges"
                  + " is removed first so the written block becomes authoritative on that"
                  + " coverage."),
          descriptor(
              MutationAction.ClearConditionalFormatting.class,
              "CLEAR_CONDITIONAL_FORMATTING",
              "Remove conditional-formatting blocks from the selected ranges on one sheet."
                  + " SELECTED removes whole blocks whose stored ranges intersect the supplied"
                  + " ranges; ALL clears every conditional-formatting block on the sheet."),
          descriptor(
              MutationAction.SetAutofilter.class,
              "SET_AUTOFILTER",
              "Create or replace one sheet-level autofilter range."
                  + " The range must include a nonblank header row and must not overlap"
                  + " an existing table range."
                  + " criteria and sortState are optional and, when supplied, are authored"
                  + " authoritatively alongside the range.",
              "criteria",
              "sortState"),
          descriptor(
              MutationAction.ClearAutofilter.class,
              "CLEAR_AUTOFILTER",
              "Clear the sheet-level autofilter range on one sheet."
                  + " Table-owned autofilters remain attached to their tables."),
          descriptor(
              MutationAction.SetTable.class,
              "SET_TABLE",
              "Create or replace one workbook-global table definition."
                  + " Table names are workbook-global and case-insensitive."
                  + " Header cells must be nonblank and unique (case-insensitive)."
                  + " Overlapping existing tables are rejected."
                  + " Any overlapping sheet-level autofilter is cleared so the table-owned"
                  + " autofilter becomes authoritative on that range."
                  + " The contract also supports advanced table metadata such as autofilter"
                  + " presence, comment, published and insert-row flags, cell-style ids,"
                  + " and per-column unique names, totals metadata, and calculated formulas."),
          descriptor(
              MutationAction.DeleteTable.class,
              "DELETE_TABLE",
              "Delete one existing workbook-global table by name and expected sheet name."),
          descriptor(
              MutationAction.DeletePivotTable.class,
              "DELETE_PIVOT_TABLE",
              "Delete one existing workbook-global pivot table by name and expected sheet name."
                  + " The expected sheet guards against accidentally deleting a same-named pivot"
                  + " after unrelated workbook changes."),
          descriptor(
              MutationAction.SetNamedRange.class,
              "SET_NAMED_RANGE",
              "Create or replace one workbook- or sheet-scoped named range."
                  + " target can be an explicit sheet plus A1 range, or a formula-defined target."),
          descriptor(
              MutationAction.DeleteNamedRange.class,
              "DELETE_NAMED_RANGE",
              "Delete one existing workbook- or sheet-scoped named range."),
          descriptor(
              MutationAction.AppendRow.class,
              "APPEND_ROW",
              "Append a row of typed values after the last value-bearing row."
                  + " Blank rows that carry only style, comment, or hyperlink metadata do not"
                  + " affect the append position."),
          descriptor(
              MutationAction.AutoSizeColumns.class,
              "AUTO_SIZE_COLUMNS",
              "Size columns deterministically from displayed cell content so the resulting"
                  + " widths are stable in headless and container runs."));
  private static final List<CatalogTypeDescriptor> ASSERTION_TYPES =
      List.of(
          descriptor(
              Assertion.Present.class,
              "EXPECT_PRESENT",
              "Require the selected target family to resolve to one or more workbook entities."),
          descriptor(
              Assertion.Absent.class,
              "EXPECT_ABSENT",
              "Require the selected target family to resolve to no workbook entities."),
          descriptor(
              Assertion.CellValue.class,
              "EXPECT_CELL_VALUE",
              "Require every selected cell to have the exact effective value."),
          descriptor(
              Assertion.DisplayValue.class,
              "EXPECT_DISPLAY_VALUE",
              "Require every selected cell to have the exact formatted display string."),
          descriptor(
              Assertion.FormulaText.class,
              "EXPECT_FORMULA_TEXT",
              "Require every selected cell to store the exact formula text."),
          descriptor(
              Assertion.CellStyle.class,
              "EXPECT_CELL_STYLE",
              "Require every selected cell to have the exact style snapshot."),
          descriptor(
              Assertion.WorkbookProtectionFacts.class,
              "EXPECT_WORKBOOK_PROTECTION",
              "Require the workbook protection report to match exactly."),
          descriptor(
              Assertion.SheetStructureFacts.class,
              "EXPECT_SHEET_STRUCTURE",
              "Require the selected sheet summary report to match exactly."),
          descriptor(
              Assertion.NamedRangeFacts.class,
              "EXPECT_NAMED_RANGE_FACTS",
              "Require the selected named-range reports to match exactly and in order."),
          descriptor(
              Assertion.TableFacts.class,
              "EXPECT_TABLE_FACTS",
              "Require the selected table reports to match exactly and in order."),
          descriptor(
              Assertion.PivotTableFacts.class,
              "EXPECT_PIVOT_TABLE_FACTS",
              "Require the selected pivot-table reports to match exactly and in order."),
          descriptor(
              Assertion.ChartFacts.class,
              "EXPECT_CHART_FACTS",
              "Require the selected chart reports to match exactly and in order."),
          descriptor(
              Assertion.AnalysisMaxSeverity.class,
              "EXPECT_ANALYSIS_MAX_SEVERITY",
              "Run one analysis query against the selected target and require its highest finding"
                  + " severity to be no higher than maximumSeverity."),
          descriptor(
              Assertion.AnalysisFindingPresent.class,
              "EXPECT_ANALYSIS_FINDING_PRESENT",
              "Run one analysis query against the selected target and require at least one"
                  + " matching finding. severity and messageContains are optional match"
                  + " refinements.",
              "severity",
              "messageContains"),
          descriptor(
              Assertion.AnalysisFindingAbsent.class,
              "EXPECT_ANALYSIS_FINDING_ABSENT",
              "Run one analysis query against the selected target and require no matching"
                  + " finding. severity and messageContains are optional match refinements.",
              "severity",
              "messageContains"),
          descriptor(
              Assertion.AllOf.class,
              "ALL_OF",
              "Require every nested assertion to pass against the same step target."),
          descriptor(
              Assertion.AnyOf.class,
              "ANY_OF",
              "Require at least one nested assertion to pass against the same step target."),
          descriptor(
              Assertion.Not.class,
              "NOT",
              "Invert one nested assertion against the same step target."));
  private static final List<CatalogTypeDescriptor> INSPECTION_QUERY_TYPES =
      List.of(
          descriptor(
              InspectionQuery.GetWorkbookSummary.class,
              "GET_WORKBOOK_SUMMARY",
              "Return workbook-level summary facts including sheet count, sheet names,"
                  + " named range count, and formula recalculation flag."
                  + " Empty workbooks return workbook.kind=EMPTY;"
                  + " non-empty workbooks return workbook.kind=WITH_SHEETS with activeSheetName"
                  + " and selectedSheetNames."),
          descriptor(
              InspectionQuery.GetPackageSecurity.class,
              "GET_PACKAGE_SECURITY",
              "Return OOXML package-encryption facts and package-signature validation results."
                  + " This reports the currently loaded workbook package state;"
                  + " after in-memory mutations, previously valid source signatures read back as"
                  + " INVALIDATED_BY_MUTATION until the saved output is re-signed."),
          descriptor(
              InspectionQuery.GetWorkbookProtection.class,
              "GET_WORKBOOK_PROTECTION",
              "Return workbook-level protection facts including structure, windows, and"
                  + " revisions lock state plus whether password hashes are present."),
          descriptor(
              InspectionQuery.GetCustomXmlMappings.class,
              "GET_CUSTOM_XML_MAPPINGS",
              "Return workbook custom-XML mapping metadata, including map identifiers,"
                  + " schema metadata, linked single cells, linked tables, and optional data"
                  + " binding facts."),
          descriptor(
              InspectionQuery.ExportCustomXmlMapping.class,
              "EXPORT_CUSTOM_XML_MAPPING",
              "Export one existing workbook custom-XML mapping as serialized XML."
                  + " mapping locates one existing map by mapId and/or name;"
                  + " validateSchema defaults to false and encoding defaults to UTF-8 when"
                  + " omitted.",
              "mapping",
              "validateSchema",
              "encoding"),
          descriptor(
              InspectionQuery.GetNamedRanges.class,
              "GET_NAMED_RANGES",
              "Return named ranges matched by the supplied selection."),
          descriptor(
              InspectionQuery.GetSheetSummary.class,
              "GET_SHEET_SUMMARY",
              "Return structural summary facts for one sheet."
                  + " Includes visibility and sheet protection state."
                  + " physicalRowCount is the number of physically materialized rows (sparse)."
                  + " lastRowIndex is the 0-based index of the last materialized row"
                  + " (-1 when empty), including metadata-only rows."
                  + " lastColumnIndex is the 0-based index of the last materialized column"
                  + " in any row (-1 when empty)."),
          descriptor(
              InspectionQuery.GetCells.class,
              "GET_CELLS",
              "Return exact cell snapshots for explicit addresses."
                  + " An invalid or out-of-range address returns INVALID_CELL_ADDRESS, not a blank."
                  + " Each snapshot includes address, declaredType, effectiveType, displayValue,"
                  + " style, and metadata fields. Type-specific fields: stringValue (STRING),"
                  + " numberValue (NUMBER), booleanValue (BOOLEAN), errorValue (ERROR),"
                  + " formula and evaluation (FORMULA). For FORMULA cells, effectiveType is FORMULA"
                  + " and the evaluated result type is in evaluation.effectiveType."
                  + " style.font.fontHeight is a plain object with both twips and points fields,"
                  + " nested under the style.font group."),
          descriptor(
              InspectionQuery.GetWindow.class,
              "GET_WINDOW",
              "Return a rectangular window of cell snapshots."
                  + " rowCount * columnCount must not exceed "
                  + InspectionQuery.MAX_WINDOW_CELLS
                  + ". Each cell snapshot has the same shape as GET_CELLS: address,"
                  + " declaredType, effectiveType, displayValue, style, metadata,"
                  + " and type-specific value fields."
                  + " For FORMULA cells, effectiveType is FORMULA and the evaluated result type"
                  + " is in evaluation.effectiveType."
                  + " Response shape: { \"window\": { \"rows\": [ { \"cells\": [...] } ] } }."
                  + " Note: the top-level key is \"window\" and cells are nested under"
                  + " window.rows[N].cells, unlike GET_CELLS which places cells directly"
                  + " under the top-level \"cells\" key."),
          descriptor(
              InspectionQuery.GetMergedRegions.class,
              "GET_MERGED_REGIONS",
              "Return the merged regions defined on one sheet."),
          descriptor(
              InspectionQuery.GetHyperlinks.class,
              "GET_HYPERLINKS",
              "Return hyperlink metadata for selected cells in the same discriminated shape used"
                  + " by SET_HYPERLINK targets. FILE targets are returned as normalized path"
                  + " strings, not file: URIs."),
          descriptor(
              InspectionQuery.GetComments.class,
              "GET_COMMENTS",
              "Return comment metadata for selected cells."),
          descriptor(
              InspectionQuery.GetDrawingObjects.class,
              "GET_DRAWING_OBJECTS",
              "Return factual drawing-object metadata for one sheet."
                  + " Read families include pictures, signature lines, simple shapes,"
                  + " connectors, groups, charts, graphic frames, and embedded objects with"
                  + " truthful anchor and package facts."),
          descriptor(
              InspectionQuery.GetCharts.class,
              "GET_CHARTS",
              "Return factual chart metadata for one sheet."
                  + " Supported authored chart families are modeled authoritatively;"
                  + " unsupported plot families or unsupported loaded detail are surfaced as"
                  + " explicit UNSUPPORTED entries with preserved plot-type tokens."),
          descriptor(
              InspectionQuery.GetArrayFormulas.class,
              "GET_ARRAY_FORMULAS",
              "Return factual array-formula group metadata for the selected sheets,"
                  + " including the stored range, top-left anchor cell, normalized formula text,"
                  + " and whether the group is single-cell."),
          descriptor(
              InspectionQuery.GetPivotTables.class,
              "GET_PIVOT_TABLES",
              "Return factual pivot-table metadata selected by workbook-global pivot-table name"
                  + " or ALL."
                  + " Supported pivots surface source, anchor, row or column labels,"
                  + " report filters, data fields, and values-axis placement."
                  + " Unsupported or malformed pivots are returned explicitly with preserved"
                  + " detail instead of causing read failure."),
          descriptor(
              InspectionQuery.GetDrawingObjectPayload.class,
              "GET_DRAWING_OBJECT_PAYLOAD",
              "Return the extracted binary payload for one existing named picture or embedded"
                  + " object."
                  + " Non-binary drawing shapes such as connectors and simple shapes are"
                  + " rejected."),
          descriptor(
              InspectionQuery.GetSheetLayout.class,
              "GET_SHEET_LAYOUT",
              GridGrindContractText.sheetLayoutReadSummary()),
          descriptor(
              InspectionQuery.GetPrintLayout.class,
              "GET_PRINT_LAYOUT",
              "Return supported print-layout metadata for one sheet, including print area,"
                  + " orientation, scaling, repeating rows or columns, and plain header or"
                  + " footer text."),
          descriptor(
              InspectionQuery.GetDataValidations.class,
              "GET_DATA_VALIDATIONS",
              "Return factual data-validation structures for the selected sheet ranges."
                  + " Supported rules include explicit lists, formula lists, comparison rules,"
                  + " and custom formulas; unsupported rules are surfaced explicitly with typed"
                  + " detail."),
          descriptor(
              InspectionQuery.GetConditionalFormatting.class,
              "GET_CONDITIONAL_FORMATTING",
              "Return factual conditional-formatting blocks for the selected sheet ranges."
                  + " Read families include authored formula and cell-value rules plus loaded"
                  + " color scales, data bars, icon sets, and explicitly unsupported rules."
                  + " Each block preserves its stored ordered ranges and rule priority data."),
          descriptor(
              InspectionQuery.GetAutofilters.class,
              "GET_AUTOFILTERS",
              "Return sheet- and table-owned autofilter metadata for one sheet."
                  + " Each entry is typed as SHEET or TABLE so ownership is explicit."),
          descriptor(
              InspectionQuery.GetTables.class,
              "GET_TABLES",
              "Return factual table metadata selected by workbook-global table name or ALL."
                  + " Each table includes range, header and totals row counts, column names,"
                  + " style metadata, and whether a table-owned autofilter is present."),
          descriptor(
              InspectionQuery.GetFormulaSurface.class,
              "GET_FORMULA_SURFACE",
              GridGrindContractText.formulaSurfaceReadSummary()),
          descriptor(
              InspectionQuery.GetSheetSchema.class,
              "GET_SHEET_SCHEMA",
              "Infer a simple schema from a rectangular sheet window."
                  + " rowCount * columnCount must not exceed "
                  + InspectionQuery.MAX_WINDOW_CELLS
                  + "."
                  + " The first row is treated as the header; dataRowCount is 0 when all header"
                  + " cells are blank."
                  + " dominantType is null when all data cells are blank, or when two or more types"
                  + " tie for the highest count."
                  + " Formula cells contribute their evaluated result type (NUMBER, STRING, etc.)"
                  + " to dominantType and observedTypes, not FORMULA."),
          descriptor(
              InspectionQuery.GetNamedRangeSurface.class,
              "GET_NAMED_RANGE_SURFACE",
              GridGrindContractText.namedRangeSurfaceReadSummary()),
          descriptor(
              InspectionQuery.AnalyzeFormulaHealth.class,
              "ANALYZE_FORMULA_HEALTH",
              GridGrindContractText.formulaHealthReadSummary()),
          descriptor(
              InspectionQuery.AnalyzeDataValidationHealth.class,
              "ANALYZE_DATA_VALIDATION_HEALTH",
              "Return analysis.checkedValidationCount, a severity summary,"
                  + " and findings such as unsupported, overlapping, or"
                  + " broken-formula rules."),
          descriptor(
              InspectionQuery.AnalyzeConditionalFormattingHealth.class,
              "ANALYZE_CONDITIONAL_FORMATTING_HEALTH",
              "Return analysis.checkedConditionalFormattingBlockCount,"
                  + " a severity summary, and conditional-formatting findings such as broken formulas,"
                  + " unsupported loaded rules, invalid target ranges, or priority collisions."),
          descriptor(
              InspectionQuery.AnalyzeAutofilterHealth.class,
              "ANALYZE_AUTOFILTER_HEALTH",
              "Return analysis.checkedAutofilterCount, a severity summary,"
                  + " and autofilter findings such as invalid ranges, blank header rows,"
                  + " or ownership mismatches between sheet-level filters and tables."),
          descriptor(
              InspectionQuery.AnalyzeTableHealth.class,
              "ANALYZE_TABLE_HEALTH",
              "Return analysis.checkedTableCount, a severity summary,"
                  + " and table findings such as overlaps, broken references,"
                  + " blank or duplicate headers, and unresolved styles."),
          descriptor(
              InspectionQuery.AnalyzePivotTableHealth.class,
              "ANALYZE_PIVOT_TABLE_HEALTH",
              "Return analysis.checkedPivotTableCount, a severity summary,"
                  + " and pivot-table findings such as missing cache parts, broken sources,"
                  + " duplicate or synthetic names, and unsupported persisted detail."),
          descriptor(
              InspectionQuery.AnalyzeHyperlinkHealth.class,
              "ANALYZE_HYPERLINK_HEALTH",
              "Return analysis.checkedHyperlinkCount, a severity summary,"
                  + " and hyperlink findings such as malformed, missing, unresolved, or broken"
                  + " targets."
                  + " Relative FILE targets are resolved against the workbook's persisted path"
                  + " when one exists."),
          descriptor(
              InspectionQuery.AnalyzeNamedRangeHealth.class,
              "ANALYZE_NAMED_RANGE_HEALTH",
              GridGrindContractText.namedRangeHealthReadSummary()),
          descriptor(
              InspectionQuery.AnalyzeWorkbookFindings.class,
              "ANALYZE_WORKBOOK_FINDINGS",
              GridGrindContractText.workbookFindingsReadSummary()));
  private static final List<CatalogNestedTypeDescriptor> NESTED_TYPE_GROUPS =
      GridGrindProtocolCatalogFieldGroupSupport.NESTED_TYPE_GROUPS;
  private static final List<CatalogPlainTypeDescriptor> PLAIN_TYPE_DESCRIPTORS =
      GridGrindProtocolCatalogFieldGroupSupport.PLAIN_TYPE_DESCRIPTORS;

  private static final Catalog CATALOG = buildCatalog();

  private GridGrindProtocolCatalog() {}

  /** Returns the minimal successful request emitted by the CLI template command. */
  public static WorkbookPlan requestTemplate() {
    return REQUEST_TEMPLATE;
  }

  /** Returns the machine-readable protocol catalog emitted by the CLI discovery command. */
  public static Catalog catalog() {
    return CATALOG;
  }

  /** Returns the ordered built-in example set exposed by the CLI and catalog. */
  public static List<GridGrindShippedExamples.ShippedExample> shippedExamples() {
    return GridGrindShippedExamples.examples();
  }

  /** Returns one built-in example by its stable discovery id, or empty when unknown. */
  public static Optional<GridGrindShippedExamples.ShippedExample> exampleFor(String id) {
    return GridGrindShippedExamples.find(id);
  }

  /**
   * Returns the single catalog entry matching the given lookup token.
   *
   * <p>Unqualified lookups succeed only when exactly one entry with that id exists across the full
   * catalog surface. When an id repeats across groups, callers must qualify it as {@code
   * <group>:<id>}; ambiguous or unknown lookups return empty.
   */
  public static Optional<TypeEntry> entryFor(String idOrQualifiedId) {
    return GridGrindProtocolCatalogLookupSupport.entryFor(CATALOG, idOrQualifiedId);
  }

  /**
   * Returns the single catalog item matching the given lookup token.
   *
   * <p>Lookups may resolve either one concrete type entry such as {@code SET_CELL} or one nested /
   * plain type-group descriptor such as {@code cellInputTypes} or {@code plainTypes:
   * chartInputType}. Ambiguous or unknown lookups return empty.
   */
  public static Optional<Object> lookupValueFor(String idOrQualifiedId) {
    return GridGrindProtocolCatalogLookupSupport.lookupValueFor(CATALOG, idOrQualifiedId);
  }

  /**
   * Returns every catalog match for the given lookup token as stable qualified ids.
   *
   * <p>Unqualified duplicate ids expand to every matching {@code <group>:<id>} so callers can
   * surface deterministic disambiguation guidance.
   */
  public static List<String> matchingEntryIds(String idOrQualifiedId) {
    return GridGrindProtocolCatalogLookupSupport.matchingEntryIds(CATALOG, idOrQualifiedId);
  }

  /**
   * Returns every catalog match for the given lookup token as stable qualified ids.
   *
   * <p>This superset includes nested and plain type-group descriptors so CLI lookup can expose the
   * exact authoring groups operators need when the protocol catalog advertises them.
   */
  public static List<String> matchingLookupIds(String idOrQualifiedId) {
    return GridGrindProtocolCatalogLookupSupport.matchingLookupIds(CATALOG, idOrQualifiedId);
  }

  /** Performs case-insensitive discovery across ids, qualified ids, groups, and summaries. */
  public static CatalogSearchResult searchCatalog(String query) {
    return GridGrindProtocolCatalogLookupSupport.search(CATALOG, query);
  }

  private static Catalog buildCatalog() {
    validateFieldShapeGroupMappings();
    validateCoverage(WorkbookPlan.WorkbookSource.class, SOURCE_TYPES);
    validateCoverage(WorkbookPlan.WorkbookPersistence.class, PERSISTENCE_TYPES);
    validateCoverage(WorkbookStep.class, STEP_TYPES);
    validateCoverage(MutationAction.class, MUTATION_ACTION_TYPES);
    validateCoverage(Assertion.class, ASSERTION_TYPES);
    validateCoverage(InspectionQuery.class, INSPECTION_QUERY_TYPES);
    for (CatalogNestedTypeDescriptor nestedTypeGroup : NESTED_TYPE_GROUPS) {
      validateCoverage(nestedTypeGroup.sealedType(), nestedTypeGroup.typeDescriptors());
    }
    return new Catalog(
        GridGrindProtocolVersion.current(),
        DISCRIMINATOR_FIELD,
        REQUEST_TYPE,
        CLI_SURFACE,
        GridGrindShippedExamples.catalogEntries(),
        publicEntries(SOURCE_TYPES),
        publicEntries(PERSISTENCE_TYPES),
        publicEntries(STEP_TYPES),
        publicEntries(MUTATION_ACTION_TYPES),
        publicEntries(ASSERTION_TYPES),
        publicEntries(INSPECTION_QUERY_TYPES),
        NESTED_TYPE_GROUPS.stream().map(GridGrindProtocolCatalog::publicGroup).toList(),
        PLAIN_TYPE_DESCRIPTORS.stream().map(GridGrindProtocolCatalog::publicPlainGroup).toList());
  }

  private static NestedTypeGroup publicGroup(CatalogNestedTypeDescriptor descriptor) {
    return new NestedTypeGroup(
        descriptor.group(),
        descriptor.discriminatorField(),
        publicEntries(descriptor.typeDescriptors()));
  }

  private static PlainTypeGroup publicPlainGroup(CatalogPlainTypeDescriptor descriptor) {
    return new PlainTypeGroup(descriptor.group(), descriptor.typeEntry());
  }

  private static List<TypeEntry> publicEntries(List<CatalogTypeDescriptor> descriptors) {
    return descriptors.stream().map(CatalogTypeDescriptor::typeEntry).toList();
  }

  static CatalogNestedTypeDescriptor nestedTypeGroup(
      String group, Class<?> sealedType, List<CatalogTypeDescriptor> typeDescriptors) {
    return new CatalogNestedTypeDescriptor(
        group, discriminatorFieldFor(sealedType), sealedType, typeDescriptors);
  }

  static CatalogPlainTypeDescriptor plainTypeDescriptor(
      String group,
      Class<? extends Record> recordType,
      String id,
      String summary,
      List<String> optionalFields) {
    return new CatalogPlainTypeDescriptor(
        group, recordType, typeEntry(recordType, id, summary, optionalFields));
  }

  static CatalogTypeDescriptor descriptor(
      Class<? extends Record> recordType, String id, String summary, String... optionalFields) {
    return new CatalogTypeDescriptor(
        recordType, typeEntry(recordType, id, summary, List.of(optionalFields)));
  }

  private static TypeEntry typeEntry(
      Class<? extends Record> recordType, String id, String summary, List<String> optionalFields) {
    WorkbookStepTargeting.TargetSurface targetSurface = targetSurfaceFor(recordType);
    return new TypeEntry(
        canonicalTypeId(recordType, id),
        summary,
        fieldEntries(recordType, optionalFields),
        targetSelectorEntries(targetSurface),
        targetSurface == null ? null : targetSurface.rule());
  }

  private static List<TargetSelectorEntry> targetSelectorEntries(
      WorkbookStepTargeting.TargetSurface targetSurface) {
    if (targetSurface == null) {
      return List.of();
    }
    return targetSurface.selectorFamilies().stream()
        .map(familyInfo -> new TargetSelectorEntry(familyInfo.family(), familyInfo.typeIds()))
        .toList();
  }

  @SuppressWarnings("unchecked")
  private static WorkbookStepTargeting.TargetSurface targetSurfaceFor(
      Class<? extends Record> recordType) {
    if (MutationAction.class.isAssignableFrom(recordType)) {
      return WorkbookStepTargeting.forMutationActionType(
          (Class<? extends MutationAction>) recordType);
    }
    if (Assertion.class.isAssignableFrom(recordType)) {
      return WorkbookStepTargeting.forAssertionType((Class<? extends Assertion>) recordType);
    }
    if (InspectionQuery.class.isAssignableFrom(recordType)) {
      return WorkbookStepTargeting.forInspectionQueryType(
          (Class<? extends InspectionQuery>) recordType);
    }
    return null;
  }

  @SuppressWarnings("unchecked")
  private static String canonicalTypeId(Class<? extends Record> recordType, String suppliedId) {
    if (WorkbookPlan.WorkbookSource.class.isAssignableFrom(recordType)) {
      return requireMatchingCatalogId(
          suppliedId,
          GridGrindProtocolTypeNames.workbookSourceTypeName(
              (Class<? extends WorkbookPlan.WorkbookSource>) recordType),
          recordType);
    }
    if (WorkbookPlan.WorkbookPersistence.class.isAssignableFrom(recordType)) {
      return requireMatchingCatalogId(
          suppliedId,
          GridGrindProtocolTypeNames.workbookPersistenceTypeName(
              (Class<? extends WorkbookPlan.WorkbookPersistence>) recordType),
          recordType);
    }
    if (WorkbookStep.class.isAssignableFrom(recordType)) {
      return requireMatchingCatalogId(
          suppliedId,
          GridGrindProtocolTypeNames.workbookStepTypeName(
              (Class<? extends WorkbookStep>) recordType),
          recordType);
    }
    if (MutationAction.class.isAssignableFrom(recordType)) {
      return requireMatchingCatalogId(
          suppliedId,
          GridGrindProtocolTypeNames.mutationActionTypeName(
              (Class<? extends MutationAction>) recordType),
          recordType);
    }
    if (Assertion.class.isAssignableFrom(recordType)) {
      return requireMatchingCatalogId(
          suppliedId,
          GridGrindProtocolTypeNames.assertionTypeName((Class<? extends Assertion>) recordType),
          recordType);
    }
    if (InspectionQuery.class.isAssignableFrom(recordType)) {
      return requireMatchingCatalogId(
          suppliedId,
          GridGrindProtocolTypeNames.inspectionQueryTypeName(
              (Class<? extends InspectionQuery>) recordType),
          recordType);
    }
    return suppliedId;
  }

  static String requireMatchingCatalogId(
      String suppliedId, String canonicalId, Class<? extends Record> recordType) {
    if (!suppliedId.equals(canonicalId)) {
      throw new IllegalStateException(
          "Catalog type id mismatch for "
              + recordType.getName()
              + ": supplied="
              + suppliedId
              + ", canonical="
              + canonicalId);
    }
    return canonicalId;
  }

  private static List<FieldEntry> fieldEntries(
      Class<? extends Record> recordType, List<String> optionalFields) {
    requiredFields(recordType, optionalFields);
    Set<String> optionalFieldSet = Set.copyOf(optionalFields);
    return Arrays.stream(recordType.getRecordComponents())
        .gather(CatalogGatherers.expandFieldsWithMetadata(optionalFieldSet))
        .toList();
  }

  private static void validateFieldShapeGroupMappings() {
    Set<Class<?>> descriptorNestedTypes =
        NESTED_TYPE_GROUPS.stream()
            .map(CatalogNestedTypeDescriptor::sealedType)
            .collect(java.util.stream.Collectors.toSet());
    Set<Class<?>> descriptorPlainTypes =
        PLAIN_TYPE_DESCRIPTORS.stream()
            .map(CatalogPlainTypeDescriptor::recordType)
            .collect(java.util.stream.Collectors.toSet());

    for (CatalogNestedTypeDescriptor descriptor : NESTED_TYPE_GROUPS) {
      CatalogFieldMetadataSupport.validateNestedTypeGroupMapping(
          descriptor.sealedType(), descriptor.group());
    }
    for (CatalogPlainTypeDescriptor descriptor : PLAIN_TYPE_DESCRIPTORS) {
      CatalogFieldMetadataSupport.validatePlainTypeGroupMapping(
          descriptor.recordType(), descriptor.group());
    }

    // Reverse check: every registered type must appear in a descriptor.
    validateReverseGroupMappings(descriptorNestedTypes, descriptorPlainTypes);
  }

  /**
   * Validates that every type registered in the field-shape maps appears in one of the provided
   * descriptor sets. Exposed as package-private so tests can exercise the failure paths with
   * synthetic descriptor sets that are intentionally incomplete.
   */
  static void validateReverseGroupMappings(
      Set<Class<?>> descriptorNestedTypes, Set<Class<?>> descriptorPlainTypes) {
    for (Class<?> registeredType : CatalogFieldMetadataSupport.registeredNestedTypes()) {
      if (!descriptorNestedTypes.contains(registeredType)) {
        throw new IllegalStateException(
            "Field-shape nested group map contains type with no catalog descriptor: "
                + registeredType.getName());
      }
    }
    for (Class<?> registeredType : CatalogFieldMetadataSupport.registeredPlainTypes()) {
      if (!descriptorPlainTypes.contains(registeredType)) {
        throw new IllegalStateException(
            "Field-shape plain group map contains type with no catalog descriptor: "
                + registeredType.getName());
      }
    }
  }

  /** Returns the required record fields after removing the explicitly optional ones. */
  static List<String> requiredFields(
      Class<? extends Record> recordType, List<String> optionalFields) {
    List<String> recordFields = recordFields(recordType);
    for (String optionalField : optionalFields) {
      if (!recordFields.contains(optionalField)) {
        throw new IllegalStateException(
            "Catalog optional field '%s' does not exist on %s"
                .formatted(optionalField, recordType.getName()));
      }
      RecordComponent component = componentNamed(recordType, optionalField);
      if (component.getType().isPrimitive()) {
        throw new IllegalStateException(
            "Catalog optional field '%s' on %s uses primitive component type %s"
                .formatted(
                    optionalField, recordType.getName(), component.getType().getSimpleName()));
      }
    }
    return recordFields.stream().filter(field -> !optionalFields.contains(field)).toList();
  }

  private static List<String> recordFields(Class<? extends Record> recordType) {
    return Arrays.stream(recordType.getRecordComponents()).map(RecordComponent::getName).toList();
  }

  private static RecordComponent componentNamed(
      Class<? extends Record> recordType, String fieldName) {
    return Arrays.stream(recordType.getRecordComponents())
        .filter(component -> component.getName().equals(fieldName))
        .findFirst()
        .orElseThrow();
  }

  private static void validateCoverage(
      Class<?> sealedType, List<CatalogTypeDescriptor> descriptors) {
    validateCoverage(
        sealedType,
        toOrderedMap(
            descriptors,
            CatalogTypeDescriptor::recordType,
            descriptor -> descriptor.typeEntry().id(),
            "catalog descriptor"));
  }

  /** Validates that a tagged union and the catalog expose the same ordered discriminator ids. */
  static void validateCoverage(Class<?> sealedType, Map<Class<?>, String> catalogIds) {
    if (sealedType.equals(WorkbookStep.class)) {
      Set<Class<?>> permitted =
          Arrays.stream(sealedType.getPermittedSubclasses())
              .collect(java.util.stream.Collectors.toCollection(java.util.LinkedHashSet::new));
      if (!permitted.equals(catalogIds.keySet())) {
        throw new IllegalStateException(
            "Catalog coverage mismatch for "
                + sealedType.getName()
                + ": permitted="
                + permitted
                + ", catalog="
                + catalogIds.keySet());
      }
      return;
    }
    discriminatorFieldFor(sealedType);
    JsonSubTypes jsonSubTypes = sealedType.getAnnotation(JsonSubTypes.class);
    if (jsonSubTypes == null) {
      throw new IllegalStateException(
          "Catalog coverage requires @JsonSubTypes on " + sealedType.getName());
    }

    Map<Class<?>, String> annotationIds =
        toOrderedMap(
            Arrays.asList(jsonSubTypes.value()),
            JsonSubTypes.Type::value,
            JsonSubTypes.Type::name,
            "annotation subtype");

    for (Class<?> recordType : catalogIds.keySet()) {
      if (!recordType.isRecord()) {
        throw new IllegalStateException(
            "Catalog entry %s does not target a record type".formatted(recordType));
      }
    }

    if (!annotationIds.keySet().equals(catalogIds.keySet())) {
      throw new IllegalStateException(
          "Catalog coverage mismatch for "
              + sealedType.getName()
              + ": annotated="
              + annotationIds.keySet()
              + ", catalog="
              + catalogIds.keySet());
    }

    for (Map.Entry<Class<?>, String> annotationEntry : annotationIds.entrySet()) {
      String catalogId = catalogIds.get(annotationEntry.getKey());
      if (!annotationEntry.getValue().equals(catalogId)) {
        throw new IllegalStateException(
            "Catalog id mismatch for "
                + annotationEntry.getKey().getName()
                + ": annotation="
                + annotationEntry.getValue()
                + ", catalog="
                + catalogId);
      }
    }
  }

  private static String discriminatorFieldFor(Class<?> sealedType) {
    JsonTypeInfo jsonTypeInfo = sealedType.getAnnotation(JsonTypeInfo.class);
    if (jsonTypeInfo == null || jsonTypeInfo.property().isBlank()) {
      throw new IllegalStateException(
          "Catalog coverage requires %s to declare a non-blank @JsonTypeInfo property"
              .formatted(sealedType.getName()));
    }
    return jsonTypeInfo.property();
  }

  @SuppressWarnings("PMD.UseConcurrentHashMap")
  static <T, K, V> Map<K, V> toOrderedMap(
      List<T> items, Function<T, K> keyFn, Function<T, V> valueFn, String label) {
    Map<K, V> result = new LinkedHashMap<>();
    for (T item : items) {
      K key = keyFn.apply(item);
      V value = valueFn.apply(item);
      if (result.containsKey(key)) {
        throw CatalogDuplicateFailures.duplicateEntryFailure(label, result.get(key), value);
      }
      result.put(key, value);
    }
    return result;
  }
}
