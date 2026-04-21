package dev.erst.gridgrind.contract.catalog;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.catalog.gather.CatalogDuplicateFailures;
import dev.erst.gridgrind.contract.catalog.gather.CatalogFieldMetadataSupport;
import dev.erst.gridgrind.contract.catalog.gather.CatalogGatherers;
import dev.erst.gridgrind.contract.dto.AutofilterFilterColumnInput;
import dev.erst.gridgrind.contract.dto.AutofilterFilterCriterionInput;
import dev.erst.gridgrind.contract.dto.AutofilterSortConditionInput;
import dev.erst.gridgrind.contract.dto.AutofilterSortStateInput;
import dev.erst.gridgrind.contract.dto.CalculationPolicyInput;
import dev.erst.gridgrind.contract.dto.CalculationReport;
import dev.erst.gridgrind.contract.dto.CalculationStrategyInput;
import dev.erst.gridgrind.contract.dto.CellAlignmentInput;
import dev.erst.gridgrind.contract.dto.CellAlignmentReport;
import dev.erst.gridgrind.contract.dto.CellBorderInput;
import dev.erst.gridgrind.contract.dto.CellBorderReport;
import dev.erst.gridgrind.contract.dto.CellBorderSideInput;
import dev.erst.gridgrind.contract.dto.CellBorderSideReport;
import dev.erst.gridgrind.contract.dto.CellColorReport;
import dev.erst.gridgrind.contract.dto.CellFillInput;
import dev.erst.gridgrind.contract.dto.CellFillReport;
import dev.erst.gridgrind.contract.dto.CellFontInput;
import dev.erst.gridgrind.contract.dto.CellFontReport;
import dev.erst.gridgrind.contract.dto.CellGradientFillInput;
import dev.erst.gridgrind.contract.dto.CellGradientFillReport;
import dev.erst.gridgrind.contract.dto.CellGradientStopInput;
import dev.erst.gridgrind.contract.dto.CellGradientStopReport;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.CellProtectionInput;
import dev.erst.gridgrind.contract.dto.CellProtectionReport;
import dev.erst.gridgrind.contract.dto.CellStyleInput;
import dev.erst.gridgrind.contract.dto.ChartInput;
import dev.erst.gridgrind.contract.dto.ChartReport;
import dev.erst.gridgrind.contract.dto.ColorInput;
import dev.erst.gridgrind.contract.dto.CommentAnchorInput;
import dev.erst.gridgrind.contract.dto.CommentInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingBlockInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingRuleInput;
import dev.erst.gridgrind.contract.dto.ConditionalFormattingThresholdInput;
import dev.erst.gridgrind.contract.dto.DataValidationErrorAlertInput;
import dev.erst.gridgrind.contract.dto.DataValidationInput;
import dev.erst.gridgrind.contract.dto.DataValidationPromptInput;
import dev.erst.gridgrind.contract.dto.DataValidationRuleInput;
import dev.erst.gridgrind.contract.dto.DifferentialBorderInput;
import dev.erst.gridgrind.contract.dto.DifferentialBorderSideInput;
import dev.erst.gridgrind.contract.dto.DifferentialStyleInput;
import dev.erst.gridgrind.contract.dto.DrawingAnchorInput;
import dev.erst.gridgrind.contract.dto.DrawingAnchorReport;
import dev.erst.gridgrind.contract.dto.DrawingMarkerInput;
import dev.erst.gridgrind.contract.dto.DrawingMarkerReport;
import dev.erst.gridgrind.contract.dto.EmbeddedObjectInput;
import dev.erst.gridgrind.contract.dto.ExecutionJournal;
import dev.erst.gridgrind.contract.dto.ExecutionJournalInput;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.FontHeightInput;
import dev.erst.gridgrind.contract.dto.FontHeightReport;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.FormulaExternalWorkbookInput;
import dev.erst.gridgrind.contract.dto.FormulaUdfFunctionInput;
import dev.erst.gridgrind.contract.dto.FormulaUdfToolpackInput;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.HeaderFooterTextInput;
import dev.erst.gridgrind.contract.dto.HyperlinkTarget;
import dev.erst.gridgrind.contract.dto.IgnoredErrorInput;
import dev.erst.gridgrind.contract.dto.NamedRangeScope;
import dev.erst.gridgrind.contract.dto.NamedRangeTarget;
import dev.erst.gridgrind.contract.dto.OoxmlEncryptionInput;
import dev.erst.gridgrind.contract.dto.OoxmlEncryptionReport;
import dev.erst.gridgrind.contract.dto.OoxmlOpenSecurityInput;
import dev.erst.gridgrind.contract.dto.OoxmlPackageSecurityReport;
import dev.erst.gridgrind.contract.dto.OoxmlPersistenceSecurityInput;
import dev.erst.gridgrind.contract.dto.OoxmlSignatureInput;
import dev.erst.gridgrind.contract.dto.OoxmlSignatureReport;
import dev.erst.gridgrind.contract.dto.PaneInput;
import dev.erst.gridgrind.contract.dto.PictureDataInput;
import dev.erst.gridgrind.contract.dto.PictureInput;
import dev.erst.gridgrind.contract.dto.PivotTableInput;
import dev.erst.gridgrind.contract.dto.PivotTableReport;
import dev.erst.gridgrind.contract.dto.PrintAreaInput;
import dev.erst.gridgrind.contract.dto.PrintLayoutInput;
import dev.erst.gridgrind.contract.dto.PrintMarginsInput;
import dev.erst.gridgrind.contract.dto.PrintScalingInput;
import dev.erst.gridgrind.contract.dto.PrintSetupInput;
import dev.erst.gridgrind.contract.dto.PrintTitleColumnsInput;
import dev.erst.gridgrind.contract.dto.PrintTitleRowsInput;
import dev.erst.gridgrind.contract.dto.RequestWarning;
import dev.erst.gridgrind.contract.dto.RichTextRunInput;
import dev.erst.gridgrind.contract.dto.ShapeInput;
import dev.erst.gridgrind.contract.dto.SheetCopyPosition;
import dev.erst.gridgrind.contract.dto.SheetDefaultsInput;
import dev.erst.gridgrind.contract.dto.SheetDisplayInput;
import dev.erst.gridgrind.contract.dto.SheetOutlineSummaryInput;
import dev.erst.gridgrind.contract.dto.SheetPresentationInput;
import dev.erst.gridgrind.contract.dto.SheetProtectionSettings;
import dev.erst.gridgrind.contract.dto.TableColumnInput;
import dev.erst.gridgrind.contract.dto.TableColumnReport;
import dev.erst.gridgrind.contract.dto.TableEntryReport;
import dev.erst.gridgrind.contract.dto.TableInput;
import dev.erst.gridgrind.contract.dto.TableStyleInput;
import dev.erst.gridgrind.contract.dto.TableStyleReport;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.dto.WorkbookProtectionInput;
import dev.erst.gridgrind.contract.dto.WorkbookProtectionReport;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.source.BinarySourceInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
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
import java.util.Objects;
import java.util.Optional;
import java.util.Set;
import java.util.function.Function;
import java.util.stream.Stream;

/**
 * Publishes the machine-readable GridGrind protocol surface used by CLI discovery commands.
 *
 * <p>This registry intentionally spans the full public contract surface, so PMD's import-count
 * heuristic is not a useful coupling signal here.
 */
@SuppressWarnings("PMD.ExcessiveImports")
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
                  "gridgrind --print-protocol-catalog [--operation <id>|<group>:<id>]",
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
                      "Column widthCharacters", "> 0 and <= 255 (Excel limit)."),
                  new CliSurface.DefinitionEntry(
                      "Row heightPoints", "> 0 and <= 1638.35 (Excel limit: 32767 twips)."),
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
                      "SET_CHART authors BAR, LINE, and PIE only; unsupported loaded chart"
                          + " detail is preserved on unrelated edits and rejected for"
                          + " authoritative mutation."),
                  new CliSurface.DefinitionEntry(
                      "Chart title formulas",
                      "SET_CHART title FORMULA and series.title FORMULA must resolve to one"
                          + " cell, directly or through one defined name."),
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
                  + " lookup as <group>:<id> such as cellInputTypes:FORMULA.",
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
                      "With --print-protocol-catalog, print one unique entry; qualify"
                          + " duplicate ids as <group>:<id>."),
                  new CliSurface.DefinitionEntry(
                      "--print-example <id>", "Print one built-in generated example request."),
                  new CliSurface.DefinitionEntry("--help, -h", "Print this help text."),
                  new CliSurface.DefinitionEntry(
                      "--version", "Print the GridGrind version and description."),
                  new CliSurface.DefinitionEntry(
                      "--license", "Print the GridGrind license and third-party notices."))),
          GridGrindContractText.standardInputRequiresRequestMessage());
  private static final List<TypeDescriptor> STEP_TYPES =
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
  private static final List<TypeDescriptor> SOURCE_TYPES =
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
  private static final List<TypeDescriptor> PERSISTENCE_TYPES =
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
  private static final List<TypeDescriptor> MUTATION_ACTION_TYPES =
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
                  + " heightPoints must be > 0 and <= 1638.35 (Excel row height limit: 32767 twips)."),
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
                  + " zoomPercent must be between 10 and 400 inclusive."),
          descriptor(
              MutationAction.SetSheetPresentation.class,
              "SET_SHEET_PRESENTATION",
              "Apply one authoritative supported sheet-presentation state to a sheet."
                  + " Omitted nested fields normalize to defaults or clear state."
                  + " The supported surface covers screen display flags, right-to-left mode,"
                  + " tab color, outline summary placement, default row and column sizing,"
                  + " and ignored-errors ranges."),
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
              MutationAction.SetChart.class,
              "SET_CHART",
              "Create or mutate one named simple chart on a sheet."
                  + " Supported authored families are BAR, LINE, and PIE."
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
              "Move one existing named picture, connector, simple shape, or embedded object"
                  + " to a new authored anchor."
                  + " anchor currently supports only TWO_CELL anchors."
                  + " Read-only loaded families such as groups and graphic frames are rejected."),
          descriptor(
              MutationAction.DeleteDrawingObject.class,
              "DELETE_DRAWING_OBJECT",
              "Delete one existing named drawing object from the sheet."
                  + " Package relationships for picture media and embedded-object parts are"
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
  private static final List<TypeDescriptor> ASSERTION_TYPES =
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
  private static final List<TypeDescriptor> INSPECTION_QUERY_TYPES =
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
                  + " Read families include pictures, simple shapes, connectors, groups,"
                  + " charts, graphic frames, and embedded objects with truthful anchor and package"
                  + " facts."),
          descriptor(
              InspectionQuery.GetCharts.class,
              "GET_CHARTS",
              "Return factual chart metadata for one sheet."
                  + " Supported simple BAR, LINE, and PIE charts are modeled authoritatively;"
                  + " unsupported plot families or multi-plot combinations are surfaced as"
                  + " explicit UNSUPPORTED entries with preserved plot-type tokens."),
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
  private static final List<NestedTypeDescriptor> NESTED_TYPE_GROUPS =
      List.of(
          nestedTypeGroup(
              "expectedCellValueTypes",
              dev.erst.gridgrind.contract.assertion.ExpectedCellValue.class,
              List.of(
                  descriptor(
                      dev.erst.gridgrind.contract.assertion.ExpectedCellValue.Blank.class,
                      "BLANK",
                      "Require the effective cell value to be blank."),
                  descriptor(
                      dev.erst.gridgrind.contract.assertion.ExpectedCellValue.Text.class,
                      "TEXT",
                      "Require the effective cell value to be one exact string."),
                  descriptor(
                      dev.erst.gridgrind.contract.assertion.ExpectedCellValue.NumericValue.class,
                      "NUMBER",
                      "Require the effective cell value to be one exact finite number."),
                  descriptor(
                      dev.erst.gridgrind.contract.assertion.ExpectedCellValue.BooleanValue.class,
                      "BOOLEAN",
                      "Require the effective cell value to be true or false."),
                  descriptor(
                      dev.erst.gridgrind.contract.assertion.ExpectedCellValue.ErrorValue.class,
                      "ERROR",
                      "Require the effective cell value to be one exact Excel error string."))),
          nestedTypeGroup(
              "textSourceTypes",
              TextSourceInput.class,
              List.of(
                  descriptor(
                      TextSourceInput.Inline.class,
                      "INLINE",
                      "Embed UTF-8 text directly in the request JSON."),
                  descriptor(
                      TextSourceInput.Utf8File.class,
                      "UTF8_FILE",
                      "Load UTF-8 text from one file path resolved in the execution"
                          + " environment."),
                  descriptor(
                      TextSourceInput.StandardInput.class,
                      "STANDARD_INPUT",
                      "Load UTF-8 text from the execution transport's bound standard input"
                          + " bytes."))),
          nestedTypeGroup(
              "binarySourceTypes",
              BinarySourceInput.class,
              List.of(
                  descriptor(
                      BinarySourceInput.InlineBase64.class,
                      "INLINE_BASE64",
                      "Embed base64-encoded binary content directly in the request JSON."),
                  descriptor(
                      BinarySourceInput.File.class,
                      "FILE",
                      "Load binary content from one file path resolved in the execution"
                          + " environment."),
                  descriptor(
                      BinarySourceInput.StandardInput.class,
                      "STANDARD_INPUT",
                      "Load binary content from the execution transport's bound standard input"
                          + " bytes."))),
          nestedTypeGroup(
              "namedRangeReportTypes",
              GridGrindResponse.NamedRangeReport.class,
              List.of(
                  descriptor(
                      GridGrindResponse.NamedRangeReport.RangeReport.class,
                      "RANGE",
                      "Exact named-range report that resolves to one typed workbook target."),
                  descriptor(
                      GridGrindResponse.NamedRangeReport.FormulaReport.class,
                      "FORMULA",
                      "Exact named-range report that remains formula-backed."))),
          nestedTypeGroup(
              "sheetProtectionReportTypes",
              GridGrindResponse.SheetProtectionReport.class,
              List.of(
                  descriptor(
                      GridGrindResponse.SheetProtectionReport.Unprotected.class,
                      "UNPROTECTED",
                      "Expect the sheet to have no protection."),
                  descriptor(
                      GridGrindResponse.SheetProtectionReport.Protected.class,
                      "PROTECTED",
                      "Expect the sheet to be protected with explicit lock settings."))),
          nestedTypeGroup(
              "tableStyleReportTypes",
              TableStyleReport.class,
              List.of(
                  descriptor(
                      TableStyleReport.None.class,
                      "NONE",
                      "Expect the table to carry no persisted style."),
                  descriptor(
                      TableStyleReport.Named.class,
                      "NAMED",
                      "Expect the table to carry one named style plus stripe/emphasis flags."))),
          nestedTypeGroup(
              "pivotTableReportTypes",
              PivotTableReport.class,
              List.of(
                  descriptor(
                      PivotTableReport.Supported.class,
                      "SUPPORTED",
                      "Exact supported pivot-table report."),
                  descriptor(
                      PivotTableReport.Unsupported.class,
                      "UNSUPPORTED",
                      "Exact unsupported pivot-table report preserved from the workbook."))),
          nestedTypeGroup(
              "pivotTableReportSourceTypes",
              PivotTableReport.Source.class,
              List.of(
                  descriptor(
                      PivotTableReport.Source.Range.class,
                      "RANGE",
                      "Pivot source resolved from one sheet range."),
                  descriptor(
                      PivotTableReport.Source.NamedRange.class,
                      "NAMED_RANGE",
                      "Pivot source resolved from one named range."),
                  descriptor(
                      PivotTableReport.Source.Table.class,
                      "TABLE",
                      "Pivot source resolved from one workbook table."))),
          nestedTypeGroup(
              "chartReportTypes",
              ChartReport.class,
              List.of(
                  descriptor(ChartReport.Bar.class, "BAR", "Exact supported bar-chart report."),
                  descriptor(ChartReport.Line.class, "LINE", "Exact supported line-chart report."),
                  descriptor(ChartReport.Pie.class, "PIE", "Exact supported pie-chart report."),
                  descriptor(
                      ChartReport.Unsupported.class,
                      "UNSUPPORTED",
                      "Exact unsupported chart report preserved from the workbook."))),
          nestedTypeGroup(
              "chartTitleReportTypes",
              ChartReport.Title.class,
              List.of(
                  descriptor(ChartReport.Title.None.class, "NONE", "No title is present."),
                  descriptor(ChartReport.Title.Text.class, "TEXT", "Static title text."),
                  descriptor(
                      ChartReport.Title.Formula.class,
                      "FORMULA",
                      "Formula-backed title with cached text."))),
          nestedTypeGroup(
              "chartLegendReportTypes",
              ChartReport.Legend.class,
              List.of(
                  descriptor(ChartReport.Legend.Hidden.class, "HIDDEN", "No legend is present."),
                  descriptor(
                      ChartReport.Legend.Visible.class,
                      "VISIBLE",
                      "Visible legend at one persisted position."))),
          nestedTypeGroup(
              "chartDataSourceReportTypes",
              ChartReport.DataSource.class,
              List.of(
                  descriptor(
                      ChartReport.DataSource.StringReference.class,
                      "STRING_REFERENCE",
                      "Formula-backed string chart source plus cached values."),
                  descriptor(
                      ChartReport.DataSource.NumericReference.class,
                      "NUMERIC_REFERENCE",
                      "Formula-backed numeric chart source plus cached values."),
                  descriptor(
                      ChartReport.DataSource.StringLiteral.class,
                      "STRING_LITERAL",
                      "Literal string chart source stored directly in the chart part."),
                  descriptor(
                      ChartReport.DataSource.NumericLiteral.class,
                      "NUMERIC_LITERAL",
                      "Literal numeric chart source stored directly in the chart part."))),
          nestedTypeGroup(
              "drawingAnchorReportTypes",
              DrawingAnchorReport.class,
              List.of(
                  descriptor(
                      DrawingAnchorReport.TwoCell.class,
                      "TWO_CELL",
                      "Drawing anchor spanning one start and end marker."),
                  descriptor(
                      DrawingAnchorReport.OneCell.class,
                      "ONE_CELL",
                      "Drawing anchor with one start marker plus explicit size."),
                  descriptor(
                      DrawingAnchorReport.Absolute.class,
                      "ABSOLUTE",
                      "Drawing anchor with absolute EMU coordinates and size."))),
          nestedTypeGroup(
              "cellInputTypes",
              CellInput.class,
              List.of(
                  descriptor(CellInput.Blank.class, "BLANK", "Write an empty cell."),
                  descriptor(
                      CellInput.Text.class,
                      "TEXT",
                      "Write a string cell value. Blank text is rejected; use BLANK for empty"
                          + " cells."),
                  descriptor(
                      CellInput.RichText.class,
                      "RICH_TEXT",
                      "Write a structured string cell value with an ordered rich-text run list."
                          + " Run text concatenates to the stored plain string value and each run"
                          + " may override font attributes independently."),
                  descriptor(CellInput.Numeric.class, "NUMBER", "Write a numeric cell value."),
                  descriptor(
                      CellInput.BooleanValue.class, "BOOLEAN", "Write a boolean cell value."),
                  descriptor(
                      CellInput.Date.class,
                      "DATE",
                      "Write an ISO-8601 date value."
                          + " Stored as an Excel serial number; GET_CELLS returns declaredType=NUMBER with a formatted displayValue."),
                  descriptor(
                      CellInput.DateTime.class,
                      "DATE_TIME",
                      "Write an ISO-8601 date-time value."
                          + " Stored as an Excel serial number; GET_CELLS returns declaredType=NUMBER with a formatted displayValue."),
                  descriptor(
                      CellInput.Formula.class,
                      "FORMULA",
                      "Write an Excel formula. A leading = sign is accepted and stripped automatically; the engine stores the formula without it."))),
          nestedTypeGroup(
              "hyperlinkTargetTypes",
              HyperlinkTarget.class,
              List.of(
                  descriptor(HyperlinkTarget.Url.class, "URL", "Attach an absolute URL target."),
                  descriptor(
                      HyperlinkTarget.Email.class,
                      "EMAIL",
                      "Attach an email target without the mailto: prefix."),
                  descriptor(
                      HyperlinkTarget.File.class,
                      "FILE",
                      "Attach a local or shared file path."
                          + " Accepts plain paths or file: URIs and normalizes to a plain path"
                          + " string."),
                  descriptor(
                      HyperlinkTarget.Document.class,
                      "DOCUMENT",
                      "Attach an internal workbook target."))),
          nestedTypeGroup(
              "paneTypes",
              PaneInput.class,
              List.of(
                  descriptor(
                      PaneInput.None.class, "NONE", "Sheet has no active pane split or freeze."),
                  descriptor(
                      PaneInput.Frozen.class,
                      "FROZEN",
                      "Freeze the sheet at the provided split and visible-origin coordinates."),
                  descriptor(
                      PaneInput.Split.class,
                      "SPLIT",
                      "Apply split panes with explicit split offsets, visible origin, and active pane."))),
          nestedTypeGroup(
              "sheetCopyPositionTypes",
              SheetCopyPosition.class,
              List.of(
                  descriptor(
                      SheetCopyPosition.AppendAtEnd.class,
                      "APPEND_AT_END",
                      "Place the copied sheet after every existing sheet."),
                  descriptor(
                      SheetCopyPosition.AtIndex.class,
                      "AT_INDEX",
                      "Place the copied sheet at the requested zero-based workbook position."))),
          nestedTypeGroup(
              "workbookSelectorTypes",
              dev.erst.gridgrind.contract.selector.WorkbookSelector.class,
              List.of(
                  descriptor(
                      dev.erst.gridgrind.contract.selector.WorkbookSelector.Current.class,
                      "CURRENT",
                      "Target the workbook currently being executed."))),
          nestedTypeGroup(
              "sheetSelectorTypes",
              dev.erst.gridgrind.contract.selector.SheetSelector.class,
              List.of(
                  descriptor(
                      dev.erst.gridgrind.contract.selector.SheetSelector.All.class,
                      "ALL",
                      "Select every sheet in workbook order."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.SheetSelector.ByName.class,
                      "BY_NAME",
                      "Select one exact sheet by name."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.SheetSelector.ByNames.class,
                      "BY_NAMES",
                      "Select one or more exact sheets by ordered name list."))),
          nestedTypeGroup(
              "cellSelectorTypes",
              dev.erst.gridgrind.contract.selector.CellSelector.class,
              List.of(
                  descriptor(
                      dev.erst.gridgrind.contract.selector.CellSelector.AllUsedInSheet.class,
                      "ALL_USED_IN_SHEET",
                      "Select every physically present cell on one sheet."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.CellSelector.ByAddress.class,
                      "BY_ADDRESS",
                      "Select one exact cell on one sheet."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.CellSelector.ByAddresses.class,
                      "BY_ADDRESSES",
                      "Select one or more exact cells on one sheet."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.CellSelector.ByQualifiedAddresses.class,
                      "BY_QUALIFIED_ADDRESSES",
                      "Select one or more exact cells across one or more sheets."))),
          nestedTypeGroup(
              "rangeSelectorTypes",
              dev.erst.gridgrind.contract.selector.RangeSelector.class,
              List.of(
                  descriptor(
                      dev.erst.gridgrind.contract.selector.RangeSelector.AllOnSheet.class,
                      "ALL_ON_SHEET",
                      "Select every matching range-backed structure on one sheet."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.RangeSelector.ByRange.class,
                      "BY_RANGE",
                      "Select one exact rectangular range on one sheet."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.RangeSelector.ByRanges.class,
                      "BY_RANGES",
                      "Select one or more exact rectangular ranges on one sheet."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.RangeSelector.RectangularWindow.class,
                      "RECTANGULAR_WINDOW",
                      "Select one rectangular window anchored at one top-left cell."))),
          nestedTypeGroup(
              "rowBandSelectorTypes",
              dev.erst.gridgrind.contract.selector.RowBandSelector.class,
              List.of(
                  descriptor(
                      dev.erst.gridgrind.contract.selector.RowBandSelector.Span.class,
                      "SPAN",
                      "Select one inclusive zero-based row span on one sheet."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.RowBandSelector.Insertion.class,
                      "INSERTION",
                      "Select one row insertion point plus row count on one sheet."))),
          nestedTypeGroup(
              "columnBandSelectorTypes",
              dev.erst.gridgrind.contract.selector.ColumnBandSelector.class,
              List.of(
                  descriptor(
                      dev.erst.gridgrind.contract.selector.ColumnBandSelector.Span.class,
                      "SPAN",
                      "Select one inclusive zero-based column span on one sheet."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.ColumnBandSelector.Insertion.class,
                      "INSERTION",
                      "Select one column insertion point plus column count on one sheet."))),
          nestedTypeGroup(
              "tableSelectorTypes",
              dev.erst.gridgrind.contract.selector.TableSelector.class,
              List.of(
                  descriptor(
                      dev.erst.gridgrind.contract.selector.TableSelector.All.class,
                      "ALL",
                      "Select every table in workbook order."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.TableSelector.ByName.class,
                      "BY_NAME",
                      "Select one workbook-global table by exact name."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.TableSelector.ByNames.class,
                      "BY_NAMES",
                      "Select one or more workbook-global tables by exact name."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.TableSelector.ByNameOnSheet.class,
                      "BY_NAME_ON_SHEET",
                      "Select one workbook-global table by exact name and expected owning sheet."))),
          nestedTypeGroup(
              "tableRowSelectorTypes",
              dev.erst.gridgrind.contract.selector.TableRowSelector.class,
              List.of(
                  descriptor(
                      dev.erst.gridgrind.contract.selector.TableRowSelector.AllRows.class,
                      "ALL_ROWS",
                      "Select every logical data row in one selected table."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.TableRowSelector.ByIndex.class,
                      "BY_INDEX",
                      "Select one zero-based data row by index in one selected table."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.TableRowSelector.ByKeyCell.class,
                      "BY_KEY_CELL",
                      "Select one logical data row by matching one key-column cell value."))),
          nestedTypeGroup(
              "tableCellSelectorTypes",
              dev.erst.gridgrind.contract.selector.TableCellSelector.class,
              List.of(
                  descriptor(
                      dev.erst.gridgrind.contract.selector.TableCellSelector.ByColumnName.class,
                      "BY_COLUMN_NAME",
                      "Select one logical cell within one selected table row by column name."))),
          nestedTypeGroup(
              "namedRangeRefSelectorTypes",
              dev.erst.gridgrind.contract.selector.NamedRangeSelector.Ref.class,
              List.of(
                  descriptor(
                      dev.erst.gridgrind.contract.selector.NamedRangeSelector.ByName.class,
                      "BY_NAME",
                      "Match a named range reference across all scopes by exact name."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.NamedRangeSelector.WorkbookScope.class,
                      "WORKBOOK_SCOPE",
                      "Match one workbook-scoped named range reference by exact name."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.NamedRangeSelector.SheetScope.class,
                      "SHEET_SCOPE",
                      "Match one sheet-scoped named range reference on one sheet."))),
          nestedTypeGroup(
              "namedRangeScopeTypes",
              NamedRangeScope.class,
              List.of(
                  descriptor(NamedRangeScope.Workbook.class, "WORKBOOK", "Target workbook scope."),
                  descriptor(
                      NamedRangeScope.Sheet.class, "SHEET", "Target one specific sheet scope."))),
          nestedTypeGroup(
              "namedRangeSelectorTypes",
              dev.erst.gridgrind.contract.selector.NamedRangeSelector.class,
              List.of(
                  descriptor(
                      dev.erst.gridgrind.contract.selector.NamedRangeSelector.All.class,
                      "ALL",
                      "Select every user-facing named range."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.NamedRangeSelector.AnyOf.class,
                      "ANY_OF",
                      "Select the union of one or more explicit named-range references."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.NamedRangeSelector.ByName.class,
                      "BY_NAME",
                      "Match a named range across all scopes by exact name."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.NamedRangeSelector.ByNames.class,
                      "BY_NAMES",
                      "Match named ranges across all scopes by exact name set."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.NamedRangeSelector.WorkbookScope.class,
                      "WORKBOOK_SCOPE",
                      "Match the workbook-scoped named range with the exact name."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.NamedRangeSelector.SheetScope.class,
                      "SHEET_SCOPE",
                      "Match the sheet-scoped named range on one sheet."))),
          nestedTypeGroup(
              "drawingObjectSelectorTypes",
              dev.erst.gridgrind.contract.selector.DrawingObjectSelector.class,
              List.of(
                  descriptor(
                      dev.erst.gridgrind.contract.selector.DrawingObjectSelector.AllOnSheet.class,
                      "ALL_ON_SHEET",
                      "Select every drawing object on one sheet."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.DrawingObjectSelector.ByName.class,
                      "BY_NAME",
                      "Select one drawing object by exact sheet-local object name."))),
          nestedTypeGroup(
              "chartSelectorTypes",
              dev.erst.gridgrind.contract.selector.ChartSelector.class,
              List.of(
                  descriptor(
                      dev.erst.gridgrind.contract.selector.ChartSelector.AllOnSheet.class,
                      "ALL_ON_SHEET",
                      "Select every chart on one sheet."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.ChartSelector.ByName.class,
                      "BY_NAME",
                      "Select one chart by exact sheet-local chart name."))),
          nestedTypeGroup(
              "pivotTableSelectorTypes",
              dev.erst.gridgrind.contract.selector.PivotTableSelector.class,
              List.of(
                  descriptor(
                      dev.erst.gridgrind.contract.selector.PivotTableSelector.All.class,
                      "ALL",
                      "Select every pivot table in workbook order."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.PivotTableSelector.ByName.class,
                      "BY_NAME",
                      "Select one workbook-global pivot table by exact name."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.PivotTableSelector.ByNames.class,
                      "BY_NAMES",
                      "Select one or more workbook-global pivot tables by exact name."),
                  descriptor(
                      dev.erst.gridgrind.contract.selector.PivotTableSelector.ByNameOnSheet.class,
                      "BY_NAME_ON_SHEET",
                      "Select one workbook-global pivot table by exact name and expected owning sheet."))),
          nestedTypeGroup(
              "drawingAnchorInputTypes",
              DrawingAnchorInput.class,
              List.of(
                  descriptor(
                      DrawingAnchorInput.TwoCell.class,
                      "TWO_CELL",
                      "Authored two-cell drawing anchor with explicit start and end markers."
                          + " behavior defaults to MOVE_AND_RESIZE when omitted.",
                      "behavior"))),
          nestedTypeGroup(
              "chartInputTypes",
              ChartInput.class,
              List.of(
                  descriptor(
                      ChartInput.Bar.class,
                      "BAR",
                      "Authored simple bar chart."
                          + " title, legend, displayBlanksAs, plotOnlyVisibleCells, varyColors,"
                          + " and barDirection default when omitted.",
                      "title",
                      "legend",
                      "displayBlanksAs",
                      "plotOnlyVisibleCells",
                      "varyColors",
                      "barDirection"),
                  descriptor(
                      ChartInput.Line.class,
                      "LINE",
                      "Authored simple line chart."
                          + " title, legend, displayBlanksAs, plotOnlyVisibleCells, and"
                          + " varyColors default when omitted.",
                      "title",
                      "legend",
                      "displayBlanksAs",
                      "plotOnlyVisibleCells",
                      "varyColors"),
                  descriptor(
                      ChartInput.Pie.class,
                      "PIE",
                      "Authored simple pie chart."
                          + " title, legend, displayBlanksAs, plotOnlyVisibleCells,"
                          + " varyColors, and firstSliceAngle are optional.",
                      "title",
                      "legend",
                      "displayBlanksAs",
                      "plotOnlyVisibleCells",
                      "varyColors",
                      "firstSliceAngle"))),
          nestedTypeGroup(
              "chartTitleInputTypes",
              ChartInput.Title.class,
              List.of(
                  descriptor(
                      ChartInput.Title.None.class, "NONE", "Remove any chart or series title."),
                  descriptor(ChartInput.Title.Text.class, "TEXT", "Use one explicit static title."),
                  descriptor(
                      ChartInput.Title.Formula.class,
                      "FORMULA",
                      "Bind the chart or series title to one workbook formula that resolves"
                          + " to one cell."))),
          nestedTypeGroup(
              "chartLegendInputTypes",
              ChartInput.Legend.class,
              List.of(
                  descriptor(ChartInput.Legend.Hidden.class, "HIDDEN", "Hide the legend entirely."),
                  descriptor(
                      ChartInput.Legend.Visible.class,
                      "VISIBLE",
                      "Show the legend at one explicit position."))),
          nestedTypeGroup(
              "pivotTableSourceTypes",
              PivotTableInput.Source.class,
              List.of(
                  descriptor(
                      PivotTableInput.Source.Range.class,
                      "RANGE",
                      "Use one explicit contiguous sheet range with the header row in the first"
                          + " row."),
                  descriptor(
                      PivotTableInput.Source.NamedRange.class,
                      "NAMED_RANGE",
                      "Use one existing workbook- or sheet-scoped named range as the pivot"
                          + " source."),
                  descriptor(
                      PivotTableInput.Source.Table.class,
                      "TABLE",
                      "Use one existing workbook-global table as the pivot source."))),
          nestedTypeGroup(
              "fontHeightTypes",
              FontHeightInput.class,
              List.of(
                  descriptor(
                      FontHeightInput.Points.class,
                      "POINTS",
                      "Specify font height in point units."
                          + " Write format: {\"type\":\"POINTS\",\"points\":13}."
                          + " Read-back (GET_CELLS, GET_WINDOW): style.fontHeight is"
                          + " {\"twips\":260,\"points\":13} with both fields present,"
                          + " not this discriminated type format."),
                  descriptor(
                      FontHeightInput.Twips.class,
                      "TWIPS",
                      "Specify font height in exact twips (20 twips = 1 point)."
                          + " Write format: {\"type\":\"TWIPS\",\"twips\":260}."
                          + " Read-back returns the same plain object shape as POINTS."))),
          nestedTypeGroup(
              "dataValidationRuleTypes",
              DataValidationRuleInput.class,
              List.of(
                  descriptor(
                      DataValidationRuleInput.ExplicitList.class,
                      "EXPLICIT_LIST",
                      "Allow only one of the supplied explicit values."
                          + " An empty values array preserves Excel's explicit-empty-list state."),
                  descriptor(
                      DataValidationRuleInput.FormulaList.class,
                      "FORMULA_LIST",
                      "Allow values from a formula-driven list expression."),
                  descriptor(
                      DataValidationRuleInput.WholeNumber.class,
                      "WHOLE_NUMBER",
                      "Apply a whole-number comparison rule."
                          + " formula2 is used only for BETWEEN and NOT_BETWEEN.",
                      "formula2"),
                  descriptor(
                      DataValidationRuleInput.DecimalNumber.class,
                      "DECIMAL_NUMBER",
                      "Apply a decimal-number comparison rule."
                          + " formula2 is used only for BETWEEN and NOT_BETWEEN.",
                      "formula2"),
                  descriptor(
                      DataValidationRuleInput.DateRule.class,
                      "DATE",
                      "Apply a date comparison rule."
                          + " formula2 is used only for BETWEEN and NOT_BETWEEN.",
                      "formula2"),
                  descriptor(
                      DataValidationRuleInput.TimeRule.class,
                      "TIME",
                      "Apply a time comparison rule."
                          + " formula2 is used only for BETWEEN and NOT_BETWEEN.",
                      "formula2"),
                  descriptor(
                      DataValidationRuleInput.TextLength.class,
                      "TEXT_LENGTH",
                      "Apply a text-length comparison rule."
                          + " formula2 is used only for BETWEEN and NOT_BETWEEN.",
                      "formula2"),
                  descriptor(
                      DataValidationRuleInput.CustomFormula.class,
                      "CUSTOM_FORMULA",
                      "Allow values that satisfy a custom formula."))),
          nestedTypeGroup(
              "autofilterFilterCriterionTypes",
              AutofilterFilterCriterionInput.class,
              List.of(
                  descriptor(
                      AutofilterFilterCriterionInput.Values.class,
                      "VALUES",
                      "Retain rows whose cell values match one or more explicit values."),
                  descriptor(
                      AutofilterFilterCriterionInput.Custom.class,
                      "CUSTOM",
                      "Retain rows that satisfy one or two comparator-based custom conditions."),
                  descriptor(
                      AutofilterFilterCriterionInput.Dynamic.class,
                      "DYNAMIC",
                      "Retain rows using one dynamic-date or moving-window autofilter rule.",
                      "value",
                      "maxValue"),
                  descriptor(
                      AutofilterFilterCriterionInput.Top10.class,
                      "TOP10",
                      "Retain top or bottom N or percent values."),
                  descriptor(
                      AutofilterFilterCriterionInput.Color.class,
                      "COLOR",
                      "Retain rows matching one cell color or font color criterion."),
                  descriptor(
                      AutofilterFilterCriterionInput.Icon.class,
                      "ICON",
                      "Retain rows matching one icon-set member."))),
          nestedTypeGroup(
              "conditionalFormattingRuleTypes",
              ConditionalFormattingRuleInput.class,
              List.of(
                  descriptor(
                      ConditionalFormattingRuleInput.FormulaRule.class,
                      "FORMULA_RULE",
                      "Apply one formula-driven conditional-formatting rule."
                          + " Supply one differential style."),
                  descriptor(
                      ConditionalFormattingRuleInput.CellValueRule.class,
                      "CELL_VALUE_RULE",
                      "Apply one cell-value comparison conditional-formatting rule."
                          + " formula2 is used only for BETWEEN and NOT_BETWEEN."
                          + " Supply one differential style.",
                      "formula2"),
                  descriptor(
                      ConditionalFormattingRuleInput.ColorScaleRule.class,
                      "COLOR_SCALE_RULE",
                      "Apply one color-scale conditional-formatting rule with ordered thresholds"
                          + " and colors."),
                  descriptor(
                      ConditionalFormattingRuleInput.DataBarRule.class,
                      "DATA_BAR_RULE",
                      "Apply one data-bar conditional-formatting rule with explicit thresholds"
                          + " and widths."),
                  descriptor(
                      ConditionalFormattingRuleInput.IconSetRule.class,
                      "ICON_SET_RULE",
                      "Apply one icon-set conditional-formatting rule with authored thresholds."),
                  descriptor(
                      ConditionalFormattingRuleInput.Top10Rule.class,
                      "TOP10_RULE",
                      "Apply one top/bottom-N conditional-formatting rule with a differential"
                          + " style."))),
          nestedTypeGroup(
              "printAreaTypes",
              PrintAreaInput.class,
              List.of(
                  descriptor(
                      PrintAreaInput.None.class, "NONE", "Sheet has no explicit print area."),
                  descriptor(
                      PrintAreaInput.Range.class,
                      "RANGE",
                      "Sheet prints the provided rectangular A1-style range."))),
          nestedTypeGroup(
              "printScalingTypes",
              PrintScalingInput.class,
              List.of(
                  descriptor(
                      PrintScalingInput.Automatic.class,
                      "AUTOMATIC",
                      "Sheet uses Excel's default scaling instead of fit-to-page counts."),
                  descriptor(
                      PrintScalingInput.Fit.class,
                      "FIT",
                      "Sheet fits printed content into the provided page counts."
                          + " A value of 0 on one axis keeps that axis unconstrained."))),
          nestedTypeGroup(
              "printTitleRowsTypes",
              PrintTitleRowsInput.class,
              List.of(
                  descriptor(
                      PrintTitleRowsInput.None.class,
                      "NONE",
                      "Sheet has no repeating print-title rows."),
                  descriptor(
                      PrintTitleRowsInput.Band.class,
                      "BAND",
                      "Sheet repeats the provided inclusive zero-based row band on every printed page."))),
          nestedTypeGroup(
              "printTitleColumnsTypes",
              PrintTitleColumnsInput.class,
              List.of(
                  descriptor(
                      PrintTitleColumnsInput.None.class,
                      "NONE",
                      "Sheet has no repeating print-title columns."),
                  descriptor(
                      PrintTitleColumnsInput.Band.class,
                      "BAND",
                      "Sheet repeats the provided inclusive zero-based column band on every printed page."))),
          nestedTypeGroup(
              "tableStyleTypes",
              TableStyleInput.class,
              List.of(
                  descriptor(TableStyleInput.None.class, "NONE", "Clear table style metadata."),
                  descriptor(
                      TableStyleInput.Named.class,
                      "NAMED",
                      "Apply one named workbook table style with explicit stripe and emphasis"
                          + " flags."))),
          nestedTypeGroup(
              "calculationStrategyTypes",
              CalculationStrategyInput.class,
              List.of(
                  descriptor(
                      CalculationStrategyInput.DoNotCalculate.class,
                      "DO_NOT_CALCULATE",
                      "Skip immediate server-side formula calculation."),
                  descriptor(
                      CalculationStrategyInput.EvaluateAll.class,
                      "EVALUATE_ALL",
                      "Preflight and evaluate every formula cell after mutation steps complete."),
                  descriptor(
                      CalculationStrategyInput.EvaluateTargets.class,
                      "EVALUATE_TARGETS",
                      "Preflight and evaluate the explicit formula-cell target list only.",
                      "cells"),
                  descriptor(
                      CalculationStrategyInput.ClearCachesOnly.class,
                      "CLEAR_CACHES_ONLY",
                      "Strip persisted formula caches without running immediate evaluation."))));
  private static final List<PlainTypeDescriptor> PLAIN_TYPE_DESCRIPTORS =
      List.of(
          plainTypeDescriptor(
              "executionJournalType",
              ExecutionJournal.class,
              "ExecutionJournal",
              "Structured execution telemetry returned on every success and failure response,"
                  + " including validation, open, calculation, step, persistence, and close"
                  + " phases.",
              List.of(
                  "planId",
                  "level",
                  "source",
                  "persistence",
                  "validation",
                  "open",
                  "calculation",
                  "persistencePhase",
                  "close",
                  "steps",
                  "warnings",
                  "outcome",
                  "events")),
          plainTypeDescriptor(
              "executionJournalSourceSummaryType",
              ExecutionJournal.SourceSummary.class,
              "ExecutionJournalSourceSummary",
              "Journal summary of the authored workbook source.",
              List.of("path")),
          plainTypeDescriptor(
              "executionJournalPersistenceSummaryType",
              ExecutionJournal.PersistenceSummary.class,
              "ExecutionJournalPersistenceSummary",
              "Journal summary of the authored persistence policy.",
              List.of("path")),
          plainTypeDescriptor(
              "executionJournalPhaseType",
              ExecutionJournal.Phase.class,
              "ExecutionJournalPhase",
              "One timed execution phase with status, timestamps, and duration.",
              List.of("startedAt", "finishedAt")),
          plainTypeDescriptor(
              "executionJournalStepType",
              ExecutionJournal.Step.class,
              "ExecutionJournalStep",
              "Per-step execution telemetry with resolved targets, timing, outcome,"
                  + " and optional failure detail.",
              List.of("failure")),
          plainTypeDescriptor(
              "executionJournalTargetType",
              ExecutionJournal.Target.class,
              "ExecutionJournalTarget",
              "One canonical target label recorded inside a step journal.",
              List.of()),
          plainTypeDescriptor(
              "executionJournalFailureClassificationType",
              ExecutionJournal.FailureClassification.class,
              "ExecutionJournalFailureClassification",
              "Structured problem-code classification for one failed step.",
              List.of()),
          plainTypeDescriptor(
              "executionJournalCalculationType",
              ExecutionJournal.Calculation.class,
              "ExecutionJournalCalculation",
              "Top-level calculation preflight and execution timings for one request.",
              List.of()),
          plainTypeDescriptor(
              "executionJournalOutcomeType",
              ExecutionJournal.Outcome.class,
              "ExecutionJournalOutcome",
              "Final outcome summary for one execution journal.",
              List.of("failedStepIndex", "failedStepId", "failureCode")),
          plainTypeDescriptor(
              "executionJournalEventType",
              ExecutionJournal.Event.class,
              "ExecutionJournalEvent",
              "Fine-grained verbose execution event emitted for live CLI rendering.",
              List.of("stepIndex", "stepId")),
          plainTypeDescriptor(
              "requestWarningType",
              RequestWarning.class,
              "RequestWarning",
              "Non-fatal authored-plan warning surfaced on success and echoed inside the execution journal.",
              List.of()),
          plainTypeDescriptor(
              "executionPolicyInputType",
              ExecutionPolicyInput.class,
              "ExecutionPolicyInput",
              GridGrindContractText.executionPolicyInputSummary(),
              List.of("mode", "journal", "calculation")),
          plainTypeDescriptor(
              "calculationPolicyInputType",
              CalculationPolicyInput.class,
              "CalculationPolicyInput",
              GridGrindContractText.calculationPolicyInputSummary(),
              List.of("strategy")),
          plainTypeDescriptor(
              "executionModeInputType",
              ExecutionModeInput.class,
              "ExecutionModeInput",
              GridGrindContractText.executionModeInputSummary(),
              List.of("readMode", "writeMode")),
          plainTypeDescriptor(
              "executionJournalInputType",
              ExecutionJournalInput.class,
              "ExecutionJournalInput",
              GridGrindContractText.executionJournalInputSummary(),
              List.of("level")),
          plainTypeDescriptor(
              "calculationReportType",
              CalculationReport.class,
              "CalculationReport",
              "Structured calculation policy, preflight classification, and execution outcome"
                  + " returned on every success and failure response.",
              List.of("preflight")),
          plainTypeDescriptor(
              "calculationPreflightType",
              CalculationReport.Preflight.class,
              "CalculationPreflightReport",
              "Formula capability classification captured before server-side evaluation begins.",
              List.of()),
          plainTypeDescriptor(
              "calculationSummaryType",
              CalculationReport.Summary.class,
              "CalculationPreflightSummary",
              "Aggregate counts for evaluable, unevaluable, and unparseable formulas.",
              List.of()),
          plainTypeDescriptor(
              "formulaCapabilityType",
              CalculationReport.FormulaCapability.class,
              "FormulaCapabilityReport",
              "One classified formula-cell capability entry from calculation preflight.",
              List.of("problemCode", "message")),
          plainTypeDescriptor(
              "calculationExecutionType",
              CalculationReport.Execution.class,
              "CalculationExecutionReport",
              "Post-execution outcome for the authored calculation policy.",
              List.of("message")),
          plainTypeDescriptor(
              "formulaEnvironmentInputType",
              FormulaEnvironmentInput.class,
              "FormulaEnvironmentInput",
              "Request-scoped formula-evaluation environment covering external workbook bindings,"
                  + " missing-workbook policy, and template-backed UDF toolpacks.",
              List.of("externalWorkbooks", "missingWorkbookPolicy", "udfToolpacks")),
          plainTypeDescriptor(
              "ooxmlOpenSecurityInputType",
              OoxmlOpenSecurityInput.class,
              "OoxmlOpenSecurityInput",
              "Optional OOXML package-open settings for encrypted existing workbook sources."
                  + " password unlocks the encrypted OOXML package before GridGrind opens the"
                  + " inner .xlsx workbook.",
              List.of("password")),
          plainTypeDescriptor(
              "ooxmlPersistenceSecurityInputType",
              OoxmlPersistenceSecurityInput.class,
              "OoxmlPersistenceSecurityInput",
              "Optional OOXML package-security settings applied during persistence."
                  + " Supply encryption, signature, or both.",
              List.of("encryption", "signature")),
          plainTypeDescriptor(
              "ooxmlEncryptionInputType",
              OoxmlEncryptionInput.class,
              "OoxmlEncryptionInput",
              "OOXML package-encryption settings for workbook persistence."
                  + " mode defaults to AGILE when omitted.",
              List.of("mode")),
          plainTypeDescriptor(
              "ooxmlSignatureInputType",
              OoxmlSignatureInput.class,
              "OoxmlSignatureInput",
              "OOXML package-signing settings for workbook persistence."
                  + " pkcs12Path resolves in the current execution environment."
                  + " keyPassword defaults to keystorePassword and digestAlgorithm defaults to"
                  + " SHA256 when omitted."
                  + " alias may be omitted only when the keystore contains exactly one"
                  + " signable private-key entry.",
              List.of("keyPassword", "alias", "digestAlgorithm", "description")),
          plainTypeDescriptor(
              "ooxmlPackageSecurityReportType",
              OoxmlPackageSecurityReport.class,
              "OoxmlPackageSecurityReport",
              "Factual OOXML package-security report covering encryption and package signatures.",
              List.of()),
          plainTypeDescriptor(
              "ooxmlEncryptionReportType",
              OoxmlEncryptionReport.class,
              "OoxmlEncryptionReport",
              "Factual OOXML package-encryption report for one workbook package."
                  + " Detail fields are present only when encrypted=true.",
              List.of()),
          plainTypeDescriptor(
              "ooxmlSignatureReportType",
              OoxmlSignatureReport.class,
              "OoxmlSignatureReport",
              "Factual OOXML package-signature report for one signature part."
                  + " state reflects the currently loaded workbook package, including"
                  + " INVALIDATED_BY_MUTATION for source signatures after in-memory edits.",
              List.of()),
          plainTypeDescriptor(
              "formulaExternalWorkbookInputType",
              FormulaExternalWorkbookInput.class,
              "FormulaExternalWorkbookInput",
              "One external workbook binding keyed by the workbook name used inside formulas.",
              List.of()),
          plainTypeDescriptor(
              "formulaUdfToolpackInputType",
              FormulaUdfToolpackInput.class,
              "FormulaUdfToolpackInput",
              "One named collection of template-backed user-defined functions.",
              List.of()),
          plainTypeDescriptor(
              "formulaUdfFunctionInputType",
              FormulaUdfFunctionInput.class,
              "FormulaUdfFunctionInput",
              "One template-backed user-defined function."
                  + " formulaTemplate may reference ARG1, ARG2, and higher placeholders."
                  + " maximumArgumentCount defaults to minimumArgumentCount when omitted.",
              List.of("maximumArgumentCount")),
          plainTypeDescriptor(
              "qualifiedCellAddressType",
              dev.erst.gridgrind.contract.selector.CellSelector.QualifiedAddress.class,
              "CellSelector.QualifiedAddress",
              "One workbook-qualified cell address used by selector-based targeted cell workflows.",
              List.of()),
          plainTypeDescriptor(
              "drawingMarkerInputType",
              DrawingMarkerInput.class,
              "DrawingMarkerInput",
              "One zero-based drawing marker with explicit column, row, and in-cell offsets.",
              List.of()),
          plainTypeDescriptor(
              "chartSeriesInputType",
              ChartInput.Series.class,
              "ChartSeriesInput",
              "One authored chart series with a title plus category and value data sources.",
              List.of("title")),
          plainTypeDescriptor(
              "chartDataSourceInputType",
              ChartInput.DataSource.class,
              "ChartDataSourceInput",
              "One contiguous chart source bound by formula or defined name.",
              List.of()),
          plainTypeDescriptor(
              "pictureDataInputType",
              PictureDataInput.class,
              "PictureDataInput",
              "One picture payload with explicit format and base64-encoded binary data.",
              List.of()),
          plainTypeDescriptor(
              "pictureInputType",
              PictureInput.class,
              "PictureInput",
              "Named picture-authoring payload for SET_PICTURE.",
              List.of("description")),
          plainTypeDescriptor(
              "shapeInputType",
              ShapeInput.class,
              "ShapeInput",
              "Named simple-shape or connector authoring payload for SET_SHAPE."
                  + " kind is limited to the authored drawing shape family."
                  + " presetGeometryToken defaults to rect for SIMPLE_SHAPE when omitted."
                  + " Invalid presetGeometryToken values are rejected non-mutatingly.",
              List.of("presetGeometryToken", "text")),
          plainTypeDescriptor(
              "embeddedObjectInputType",
              EmbeddedObjectInput.class,
              "EmbeddedObjectInput",
              "Named embedded-object authoring payload for SET_EMBEDDED_OBJECT."
                  + " base64Data holds the embedded package bytes and previewImage holds the"
                  + " visible preview raster.",
              List.of()),
          plainTypeDescriptor(
              "commentInputType",
              CommentInput.class,
              "CommentInput",
              "Comment payload attached to one cell."
                  + " Comments can carry ordered rich-text runs and an explicit anchor box.",
              List.of("visible", "runs", "anchor")),
          plainTypeDescriptor(
              "commentAnchorInputType",
              CommentAnchorInput.class,
              "CommentAnchorInput",
              "Explicit comment-anchor bounds measured in zero-based column and row indexes.",
              List.of()),
          plainTypeDescriptor(
              "namedRangeTargetType",
              NamedRangeTarget.class,
              "NamedRangeTarget",
              "Named-range target payload."
                  + " Supply either sheetName plus range, or formula by itself.",
              List.of("sheetName", "range", "formula")),
          plainTypeDescriptor(
              "sheetProtectionSettingsType",
              SheetProtectionSettings.class,
              "SheetProtectionSettings",
              "Supported sheet-protection lock flags authored and reported by GridGrind.",
              List.of()),
          plainTypeDescriptor(
              "cellStyleInputType",
              CellStyleInput.class,
              "CellStyleInput",
              "Style patch applied to a cell or range; at least one field must be set."
                  + " Colors use #RRGGBB hex and style subgroups are nested explicitly.",
              List.of("numberFormat", "alignment", "font", "fill", "border", "protection")),
          plainTypeDescriptor(
              "cellAlignmentInputType",
              CellAlignmentInput.class,
              "CellAlignmentInput",
              "Alignment patch for cell styling; at least one field must be set."
                  + " textRotation uses XSSF's explicit 0-180 degree scale and indentation uses"
                  + " Excel's 0-250 cell-indent range.",
              List.of(
                  "wrapText",
                  "horizontalAlignment",
                  "verticalAlignment",
                  "textRotation",
                  "indentation")),
          plainTypeDescriptor(
              "cellFontInputType",
              CellFontInput.class,
              "CellFontInput",
              "Font patch for cell styling; at least one field must be set."
                  + " Colors can use RGB, theme, indexed, and tint semantics.",
              List.of(
                  "bold",
                  "italic",
                  "fontName",
                  "fontHeight",
                  "fontColor",
                  "fontColorTheme",
                  "fontColorIndexed",
                  "fontColorTint",
                  "underline",
                  "strikeout")),
          plainTypeDescriptor(
              "colorInputType",
              ColorInput.class,
              "ColorInput",
              "Color payload preserving RGB, theme, indexed, and tint semantics."
                  + " At least one of rgb, theme, or indexed must be supplied.",
              List.of("rgb", "theme", "indexed", "tint")),
          plainTypeDescriptor(
              "richTextRunInputType",
              RichTextRunInput.class,
              "RichTextRunInput",
              "One ordered rich-text run for a string cell."
                  + " text must be non-empty; font is an optional override patch."
                  + " The ordered run texts concatenate to the stored plain string value.",
              List.of("font")),
          plainTypeDescriptor(
              "cellFillInputType",
              CellFillInput.class,
              "CellFillInput",
              "Fill patch for cell styling. pattern controls solid and patterned fills;"
                  + " colors can use RGB, theme, indexed, and tint semantics."
                  + " gradient is mutually exclusive with patterned fill fields.",
              List.of(
                  "pattern",
                  "foregroundColor",
                  "foregroundColorTheme",
                  "foregroundColorIndexed",
                  "foregroundColorTint",
                  "backgroundColor",
                  "backgroundColorTheme",
                  "backgroundColorIndexed",
                  "backgroundColorTint",
                  "gradient")),
          plainTypeDescriptor(
              "cellGradientFillInputType",
              CellGradientFillInput.class,
              "CellGradientFillInput",
              "Gradient fill payload for cell-style authoring."
                  + " LINEAR gradients use degree, PATH gradients use left/right/top/bottom,"
                  + " and the two geometry modes must not be mixed."
                  + " stops must contain at least two entries.",
              List.of("type", "degree", "left", "right", "top", "bottom")),
          plainTypeDescriptor(
              "cellGradientStopInputType",
              CellGradientStopInput.class,
              "CellGradientStopInput",
              "One gradient stop with a normalized position between 0.0 and 1.0.",
              List.of()),
          plainTypeDescriptor(
              "cellBorderInputType",
              CellBorderInput.class,
              "CellBorderInput",
              "Border patch for cell styling; at least one side must be set."
                  + " Use 'all' as shorthand for all four sides.",
              List.of("all", "top", "right", "bottom", "left")),
          plainTypeDescriptor(
              "cellBorderSideInputType",
              CellBorderSideInput.class,
              "CellBorderSideInput",
              "One border side defined by its border style and optional color semantics.",
              List.of("style", "color", "colorTheme", "colorIndexed", "colorTint")),
          plainTypeDescriptor(
              "cellProtectionInputType",
              CellProtectionInput.class,
              "CellProtectionInput",
              "Cell protection patch; at least one field must be set."
                  + " These flags matter when sheet protection is enabled.",
              List.of("locked", "hiddenFormula")),
          plainTypeDescriptor(
              "dataValidationInputType",
              DataValidationInput.class,
              "DataValidationInput",
              "Supported data-validation definition attached to one sheet range.",
              List.of("allowBlank", "suppressDropDownArrow", "prompt", "errorAlert")),
          plainTypeDescriptor(
              "dataValidationPromptInputType",
              DataValidationPromptInput.class,
              "DataValidationPromptInput",
              "Optional prompt-box configuration shown when a validated cell is selected.",
              List.of("showPromptBox")),
          plainTypeDescriptor(
              "dataValidationErrorAlertInputType",
              DataValidationErrorAlertInput.class,
              "DataValidationErrorAlertInput",
              "Optional error-box configuration shown when invalid data is entered.",
              List.of("showErrorBox")),
          plainTypeDescriptor(
              "autofilterCustomConditionInputType",
              AutofilterFilterCriterionInput.CustomConditionInput.class,
              "AutofilterCustomConditionInput",
              "One comparator-value pair nested inside a custom autofilter criterion.",
              List.of()),
          plainTypeDescriptor(
              "autofilterFilterColumnInputType",
              AutofilterFilterColumnInput.class,
              "AutofilterFilterColumnInput",
              "One authored autofilter filter-column payload with an explicit column criterion.",
              List.of("showButton")),
          plainTypeDescriptor(
              "autofilterSortConditionInputType",
              AutofilterSortConditionInput.class,
              "AutofilterSortConditionInput",
              "One authored sort condition nested inside an autofilter sort state.",
              List.of("sortBy", "color", "iconId")),
          plainTypeDescriptor(
              "autofilterSortStateInputType",
              AutofilterSortStateInput.class,
              "AutofilterSortStateInput",
              "Authored autofilter sort-state payload with one or more ordered sort conditions.",
              List.of("caseSensitive", "columnSort", "sortMethod")),
          plainTypeDescriptor(
              "conditionalFormattingBlockInputType",
              ConditionalFormattingBlockInput.class,
              "ConditionalFormattingBlockInput",
              "One authored conditional-formatting block with ordered target ranges and rules."
                  + " rules must not be empty; ranges must be unique.",
              List.of()),
          plainTypeDescriptor(
              "conditionalFormattingThresholdInputType",
              ConditionalFormattingThresholdInput.class,
              "ConditionalFormattingThresholdInput",
              "Threshold payload shared by authored advanced conditional-formatting rules.",
              List.of("formula", "value")),
          plainTypeDescriptor(
              "headerFooterTextInputType",
              HeaderFooterTextInput.class,
              "HeaderFooterTextInput",
              "Plain left, center, and right header or footer text segments."
                  + " Null fields default to empty string.",
              List.of("left", "center", "right")),
          plainTypeDescriptor(
              "differentialStyleInputType",
              DifferentialStyleInput.class,
              "DifferentialStyleInput",
              "Differential style payload used by authored conditional-formatting rules."
                  + " At least one field must be set. Colors use #RRGGBB hex.",
              List.of(
                  "numberFormat",
                  "bold",
                  "italic",
                  "fontHeight",
                  "fontColor",
                  "underline",
                  "strikeout",
                  "fillColor",
                  "border")),
          plainTypeDescriptor(
              "differentialBorderInputType",
              DifferentialBorderInput.class,
              "DifferentialBorderInput",
              "Conditional-formatting differential border patch; at least one side must be set."
                  + " Use 'all' as shorthand for all four sides.",
              List.of("all", "top", "right", "bottom", "left")),
          plainTypeDescriptor(
              "differentialBorderSideInputType",
              DifferentialBorderSideInput.class,
              "DifferentialBorderSideInput",
              "One conditional-formatting differential border side defined by style and optional"
                  + " color.",
              List.of()),
          plainTypeDescriptor(
              "ignoredErrorInputType",
              IgnoredErrorInput.class,
              "IgnoredErrorInput",
              "One ignored-error block anchored to one A1-style range plus one or more"
                  + " ignored-error families.",
              List.of()),
          plainTypeDescriptor(
              "printLayoutInputType",
              PrintLayoutInput.class,
              "PrintLayoutInput",
              "Authoritative supported print-layout payload for one SET_PRINT_LAYOUT request."
                  + " All fields are optional and normalize to defaults when omitted.",
              List.of(
                  "printArea",
                  "orientation",
                  "scaling",
                  "repeatingRows",
                  "repeatingColumns",
                  "header",
                  "footer",
                  "setup")),
          plainTypeDescriptor(
              "printMarginsInputType",
              PrintMarginsInput.class,
              "PrintMarginsInput",
              "Explicit print margins measured in the workbook's stored inch-based values.",
              List.of()),
          plainTypeDescriptor(
              "printSetupInputType",
              PrintSetupInput.class,
              "PrintSetupInput",
              "Advanced page-setup payload nested under print-layout authoring."
                  + " All fields are optional and normalize to defaults when omitted.",
              List.of(
                  "margins",
                  "printGridlines",
                  "horizontallyCentered",
                  "verticallyCentered",
                  "paperSize",
                  "draft",
                  "blackAndWhite",
                  "copies",
                  "useFirstPageNumber",
                  "firstPageNumber",
                  "rowBreaks",
                  "columnBreaks")),
          plainTypeDescriptor(
              "sheetDefaultsInputType",
              SheetDefaultsInput.class,
              "SheetDefaultsInput",
              "Default row and column sizing authored as part of sheet-presentation state."
                  + " All fields are optional and normalize to defaults when omitted.",
              List.of("defaultColumnWidth", "defaultRowHeightPoints")),
          plainTypeDescriptor(
              "sheetDisplayInputType",
              SheetDisplayInput.class,
              "SheetDisplayInput",
              "Screen-facing sheet display flags authored as part of sheet-presentation state."
                  + " All fields are optional and normalize to defaults when omitted.",
              List.of(
                  "displayGridlines",
                  "displayZeros",
                  "displayRowColHeadings",
                  "displayFormulas",
                  "rightToLeft")),
          plainTypeDescriptor(
              "sheetOutlineSummaryInputType",
              SheetOutlineSummaryInput.class,
              "SheetOutlineSummaryInput",
              "Outline-summary placement authored as part of sheet-presentation state."
                  + " All fields are optional and normalize to defaults when omitted.",
              List.of("rowSumsBelow", "rowSumsRight")),
          plainTypeDescriptor(
              "sheetPresentationInputType",
              SheetPresentationInput.class,
              "SheetPresentationInput",
              "Authoritative sheet-presentation payload for one SET_SHEET_PRESENTATION request."
                  + " All fields are optional and normalize to defaults or clear state when"
                  + " omitted.",
              List.of("display", "tabColor", "outlineSummary", "sheetDefaults", "ignoredErrors")),
          plainTypeDescriptor(
              "pivotTableInputType",
              PivotTableInput.class,
              "PivotTableInput",
              "Workbook-global pivot-table definition for one SET_PIVOT_TABLE request."
                  + " Source-column assignments across rowLabels, columnLabels, reportFilters,"
                  + " and dataFields must be disjoint."
                  + " reportFilters require anchor.topLeftAddress on row 3 or lower.",
              List.of("rowLabels", "columnLabels", "reportFilters")),
          plainTypeDescriptor(
              "pivotTableAnchorInputType",
              PivotTableInput.Anchor.class,
              "PivotTableAnchorInput",
              "Top-left anchor for a pivot table rendered on its destination sheet."
                  + " The address must be a single-cell A1 reference.",
              List.of()),
          plainTypeDescriptor(
              "pivotTableDataFieldInputType",
              PivotTableInput.DataField.class,
              "PivotTableDataFieldInput",
              "One authored pivot data field bound to a source column and aggregation function."
                  + " displayName defaults to sourceColumnName when omitted.",
              List.of("displayName", "valueFormat")),
          plainTypeDescriptor(
              "tableColumnInputType",
              TableColumnInput.class,
              "TableColumnInput",
              "Advanced table-column metadata applied by zero-based ordinal column index.",
              List.of(
                  "uniqueName", "totalsRowLabel", "totalsRowFunction", "calculatedColumnFormula")),
          plainTypeDescriptor(
              "tableInputType",
              TableInput.class,
              "TableInput",
              "Workbook-global table definition for one SET_TABLE request.",
              List.of(
                  "showTotalsRow",
                  "hasAutofilter",
                  "comment",
                  "published",
                  "insertRow",
                  "insertRowShift",
                  "headerRowCellStyle",
                  "dataCellStyle",
                  "totalsRowCellStyle",
                  "columns")),
          plainTypeDescriptor(
              "workbookProtectionInputType",
              WorkbookProtectionInput.class,
              "WorkbookProtectionInput",
              "Workbook-protection payload covering workbook and revisions lock state plus"
                  + " optional passwords.",
              List.of(
                  "structureLocked",
                  "windowsLocked",
                  "revisionsLocked",
                  "workbookPassword",
                  "revisionsPassword")),
          plainTypeDescriptor(
              "workbookProtectionReportType",
              WorkbookProtectionReport.class,
              "WorkbookProtectionReport",
              "Exact workbook-protection report covering structure, windows, revisions,"
                  + " and password-hash presence flags.",
              List.of()),
          plainTypeDescriptor(
              "sheetSummaryReportType",
              GridGrindResponse.SheetSummaryReport.class,
              "SheetSummaryReport",
              "Exact sheet summary report including visibility, protection, and structural counts.",
              List.of()),
          plainTypeDescriptor(
              "cellStyleReportType",
              GridGrindResponse.CellStyleReport.class,
              "CellStyleReport",
              "Exact effective cell-style report used by style assertions.",
              List.of()),
          plainTypeDescriptor(
              "cellAlignmentReportType",
              CellAlignmentReport.class,
              "CellAlignmentReport",
              "Exact cell-alignment report.",
              List.of()),
          plainTypeDescriptor(
              "cellFontReportType",
              CellFontReport.class,
              "CellFontReport",
              "Exact cell-font report.",
              List.of("fontColor")),
          plainTypeDescriptor(
              "cellFillReportType",
              CellFillReport.class,
              "CellFillReport",
              "Exact cell-fill report including pattern, colors, or gradient payload.",
              List.of("foregroundColor", "backgroundColor", "gradient")),
          plainTypeDescriptor(
              "cellBorderReportType",
              CellBorderReport.class,
              "CellBorderReport",
              "Exact four-sided cell-border report.",
              List.of()),
          plainTypeDescriptor(
              "cellBorderSideReportType",
              CellBorderSideReport.class,
              "CellBorderSideReport",
              "Exact one-sided cell-border report.",
              List.of("color")),
          plainTypeDescriptor(
              "cellProtectionReportType",
              CellProtectionReport.class,
              "CellProtectionReport",
              "Exact cell-protection report.",
              List.of()),
          plainTypeDescriptor(
              "cellColorReportType",
              CellColorReport.class,
              "CellColorReport",
              "Exact workbook color report preserving RGB, theme, indexed, and tint semantics.",
              List.of("rgb", "theme", "indexed", "tint")),
          plainTypeDescriptor(
              "fontHeightReportType",
              FontHeightReport.class,
              "FontHeightReport",
              "Exact font-height report expressed in twips and points.",
              List.of()),
          plainTypeDescriptor(
              "cellGradientFillReportType",
              CellGradientFillReport.class,
              "CellGradientFillReport",
              "Exact gradient-fill report with geometry and stops.",
              List.of("degree", "left", "right", "top", "bottom")),
          plainTypeDescriptor(
              "cellGradientStopReportType",
              CellGradientStopReport.class,
              "CellGradientStopReport",
              "Exact gradient stop report.",
              List.of()),
          plainTypeDescriptor(
              "tableEntryReportType",
              TableEntryReport.class,
              "TableEntryReport",
              "Exact workbook table report used by table-facts assertions.",
              List.of("comment", "headerRowCellStyle", "dataCellStyle", "totalsRowCellStyle")),
          plainTypeDescriptor(
              "tableColumnReportType",
              TableColumnReport.class,
              "TableColumnReport",
              "Exact table-column report.",
              List.of(
                  "uniqueName", "totalsRowLabel", "totalsRowFunction", "calculatedColumnFormula")),
          plainTypeDescriptor(
              "drawingMarkerReportType",
              DrawingMarkerReport.class,
              "DrawingMarkerReport",
              "Exact cell-relative drawing marker report.",
              List.of()),
          plainTypeDescriptor(
              "pivotTableAnchorReportType",
              PivotTableReport.Anchor.class,
              "PivotTableAnchorReport",
              "Exact pivot-table anchor report.",
              List.of()),
          plainTypeDescriptor(
              "pivotTableFieldReportType",
              PivotTableReport.Field.class,
              "PivotTableFieldReport",
              "Exact pivot field report bound to one source column.",
              List.of()),
          plainTypeDescriptor(
              "pivotTableDataFieldReportType",
              PivotTableReport.DataField.class,
              "PivotTableDataFieldReport",
              "Exact pivot data-field report.",
              List.of("valueFormat")),
          plainTypeDescriptor(
              "chartAxisReportType",
              ChartReport.Axis.class,
              "ChartAxisReport",
              "Exact chart-axis report.",
              List.of()),
          plainTypeDescriptor(
              "chartSeriesReportType",
              ChartReport.Series.class,
              "ChartSeriesReport",
              "Exact chart-series report.",
              List.of()));
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
    List<CatalogEntryRef> matches = matchingEntryRefs(idOrQualifiedId);
    return matches.size() == 1 ? Optional.of(matches.getFirst().entry()) : Optional.empty();
  }

  /**
   * Returns every catalog match for the given lookup token as stable qualified ids.
   *
   * <p>Unqualified duplicate ids expand to every matching {@code <group>:<id>} so callers can
   * surface deterministic disambiguation guidance.
   */
  public static List<String> matchingEntryIds(String idOrQualifiedId) {
    return matchingEntryRefs(idOrQualifiedId).stream().map(CatalogEntryRef::qualifiedId).toList();
  }

  private static List<CatalogEntryRef> matchingEntryRefs(String idOrQualifiedId) {
    String lookup =
        CatalogRecordValidation.requireNonBlank(idOrQualifiedId, "idOrQualifiedId").trim();
    int separator = lookup.indexOf(':');
    if (separator >= 0) {
      String group = lookup.substring(0, separator).trim();
      String id = lookup.substring(separator + 1).trim();
      if (group.isEmpty() || id.isEmpty()) {
        return List.of();
      }
      return allEntryRefs().stream()
          .filter(entryRef -> entryRef.group().equals(group) && entryRef.entry().id().equals(id))
          .toList();
    }
    return allEntryRefs().stream()
        .filter(entryRef -> entryRef.entry().id().equals(lookup))
        .toList();
  }

  private static List<CatalogEntryRef> allEntryRefs() {
    return Stream.of(
            Stream.of(new CatalogEntryRef("requestType", CATALOG.requestType())),
            entryRefs("sourceTypes", CATALOG.sourceTypes()).stream(),
            entryRefs("persistenceTypes", CATALOG.persistenceTypes()).stream(),
            entryRefs("stepTypes", CATALOG.stepTypes()).stream(),
            entryRefs("mutationActionTypes", CATALOG.mutationActionTypes()).stream(),
            entryRefs("assertionTypes", CATALOG.assertionTypes()).stream(),
            entryRefs("inspectionQueryTypes", CATALOG.inspectionQueryTypes()).stream(),
            CATALOG.nestedTypes().stream()
                .flatMap(group -> entryRefs(group.group(), group.types()).stream()),
            CATALOG.plainTypes().stream()
                .map(group -> new CatalogEntryRef(group.group(), group.type())))
        .flatMap(Function.identity())
        .toList();
  }

  private static List<CatalogEntryRef> entryRefs(String group, List<TypeEntry> entries) {
    return entries.stream().map(entry -> new CatalogEntryRef(group, entry)).toList();
  }

  private static Catalog buildCatalog() {
    validateFieldShapeGroupMappings();
    validateCoverage(WorkbookPlan.WorkbookSource.class, SOURCE_TYPES);
    validateCoverage(WorkbookPlan.WorkbookPersistence.class, PERSISTENCE_TYPES);
    validateCoverage(WorkbookStep.class, STEP_TYPES);
    validateCoverage(MutationAction.class, MUTATION_ACTION_TYPES);
    validateCoverage(Assertion.class, ASSERTION_TYPES);
    validateCoverage(InspectionQuery.class, INSPECTION_QUERY_TYPES);
    for (NestedTypeDescriptor nestedTypeGroup : NESTED_TYPE_GROUPS) {
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

  private static NestedTypeGroup publicGroup(NestedTypeDescriptor descriptor) {
    return new NestedTypeGroup(
        descriptor.group(),
        descriptor.discriminatorField(),
        publicEntries(descriptor.typeDescriptors()));
  }

  private static PlainTypeGroup publicPlainGroup(PlainTypeDescriptor descriptor) {
    return new PlainTypeGroup(descriptor.group(), descriptor.typeEntry());
  }

  private static List<TypeEntry> publicEntries(List<TypeDescriptor> descriptors) {
    return descriptors.stream().map(TypeDescriptor::typeEntry).toList();
  }

  private static NestedTypeDescriptor nestedTypeGroup(
      String group, Class<?> sealedType, List<TypeDescriptor> typeDescriptors) {
    return new NestedTypeDescriptor(
        group, discriminatorFieldFor(sealedType), sealedType, typeDescriptors);
  }

  private static PlainTypeDescriptor plainTypeDescriptor(
      String group,
      Class<? extends Record> recordType,
      String id,
      String summary,
      List<String> optionalFields) {
    return new PlainTypeDescriptor(
        group, recordType, typeEntry(recordType, id, summary, optionalFields));
  }

  private static TypeDescriptor descriptor(
      Class<? extends Record> recordType, String id, String summary, String... optionalFields) {
    return new TypeDescriptor(
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
            .map(NestedTypeDescriptor::sealedType)
            .collect(java.util.stream.Collectors.toSet());
    Set<Class<?>> descriptorPlainTypes =
        PLAIN_TYPE_DESCRIPTORS.stream()
            .map(PlainTypeDescriptor::recordType)
            .collect(java.util.stream.Collectors.toSet());

    for (NestedTypeDescriptor descriptor : NESTED_TYPE_GROUPS) {
      CatalogFieldMetadataSupport.validateNestedTypeGroupMapping(
          descriptor.sealedType(), descriptor.group());
    }
    for (PlainTypeDescriptor descriptor : PLAIN_TYPE_DESCRIPTORS) {
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

  private static void validateCoverage(Class<?> sealedType, List<TypeDescriptor> descriptors) {
    validateCoverage(
        sealedType,
        toOrderedMap(
            descriptors,
            TypeDescriptor::recordType,
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

  private record TypeDescriptor(Class<? extends Record> recordType, TypeEntry typeEntry) {
    private TypeDescriptor {
      Objects.requireNonNull(recordType, "recordType must not be null");
      Objects.requireNonNull(typeEntry, "typeEntry must not be null");
    }
  }

  private record PlainTypeDescriptor(
      String group, Class<? extends Record> recordType, TypeEntry typeEntry) {
    private PlainTypeDescriptor {
      group = CatalogRecordValidation.requireNonBlank(group, "group");
      Objects.requireNonNull(recordType, "recordType must not be null");
      Objects.requireNonNull(typeEntry, "typeEntry must not be null");
    }
  }

  private record NestedTypeDescriptor(
      String group,
      String discriminatorField,
      Class<?> sealedType,
      List<TypeDescriptor> typeDescriptors) {
    private NestedTypeDescriptor {
      group = CatalogRecordValidation.requireNonBlank(group, "group");
      discriminatorField =
          CatalogRecordValidation.requireNonBlank(discriminatorField, "discriminatorField");
      Objects.requireNonNull(sealedType, "sealedType must not be null");
      typeDescriptors = List.copyOf(typeDescriptors);
      for (TypeDescriptor typeDescriptor : typeDescriptors) {
        Objects.requireNonNull(typeDescriptor, "typeDescriptors must not contain nulls");
      }
    }
  }

  private record CatalogEntryRef(String group, TypeEntry entry) {
    private CatalogEntryRef {
      group = CatalogRecordValidation.requireNonBlank(group, "group");
      Objects.requireNonNull(entry, "entry must not be null");
    }

    private String qualifiedId() {
      return group + ":" + entry.id();
    }
  }
}
