package dev.erst.gridgrind.cli.render;

import dev.erst.gridgrind.core.catalog.GridGrindInspectionContractText;
import dev.erst.gridgrind.core.catalog.GridGrindOoxmlWriteEncryptionContractText;
import dev.erst.gridgrind.core.catalog.GridGrindRequestSurfaceContractText;
import dev.erst.gridgrind.core.semantics.BackendCapabilities;
import java.util.List;

/** Synopsis-first CLI help sections. */
final class HelpSynopsis {
  private HelpSynopsis() {}

  static HelpSurfaceReport.CliSection usage() {
    return new HelpSurfaceReport.CliSection(
        "Usage",
        List.of(
            "gridgrind run --request <path|-> --grant <path> [--response <path>] [--pretty]",
            "gridgrind doctor --request <path|-> [--grant <path>] [--response <path>] [--pretty]",
            "gridgrind catalog [search <text> | show <group>:<id>] [--response <path>] [--pretty]",
            "gridgrind schema request|grant|response [--response <path>] [--pretty]",
            "gridgrind recipe list | show <ID> | match <text> [--response <path>] [--pretty]",
            "gridgrind recipe materialize <ID> --workspace <new-directory> [--pretty]",
            "gridgrind template request|grant [--response <path>] [--pretty]",
            "gridgrind help [overview|protocol|guidance] [--format text|structured]",
            "gridgrind version | license [--format text|structured]",
            "For stdin requests use --request - with --execution-root <directory>.",
            "Run and doctor also accept --temp-root and --secrets-provider paths."));
  }

  static HelpSurfaceReport.CliWorkflowSection workflows() {
    return new HelpSurfaceReport.CliWorkflowSection(
        "Workflows",
        List.of(
            new HelpSurfaceReport.WorkflowEntry(
                "Discover",
                List.of(
                    "List recipes: gridgrind recipe list",
                    "Match a goal: gridgrind recipe match \"monthly dashboard\"",
                    "Find protocol capabilities: gridgrind catalog search \"chart title\"",
                    "Inspect an exact shape: gridgrind catalog show mutationActionTypes:SET_CELL")),
            new HelpSurfaceReport.WorkflowEntry(
                "Prepare and authorize",
                List.of(
                    "Create a request: gridgrind recipe show BUDGET --response budget-request.json",
                    "Create fail-closed authority: gridgrind template grant --response grant.json",
                    "Edit the grant to authorize the request's exact resources, targets and publication.",
                    "Diagnose: gridgrind doctor --request budget-request.json --grant grant.json")),
            new HelpSurfaceReport.WorkflowEntry(
                "Execute",
                List.of(
                    "Run: gridgrind run --request budget-request.json --grant grant.json --response response.json",
                    "Materialize asset-backed recipes into a new workspace before execution.",
                    "Inspect the structured result, warnings, assertions and publication outcome."))));
  }

  static HelpSurfaceReport.CliSection execution() {
    return new HelpSurfaceReport.CliSection(
        "Execution",
        List.of(
            "Without --response, the command payload is the sole stdout content. With --response,"
                + " that payload is written to a new file; existing response files are not replaced.",
            "A response-file write failure with writable stdout recovers the already-rendered"
                + " payload there unchanged and adds one compact transport notice on stderr."
                + " GridGrind never moves a primary payload to stderr.",
            "Use --format structured for JSON help, version and license reports. Use --pretty"
                + " to indent JSON output.",
            "GridGrind executes ordered steps in sequence, then saves the workbook (unless"
                + " persistence is NONE); if any step fails, no workbook is written.",
            "source.type=NEW starts with zero sheets; ENSURE_SHEET creates the first" + " sheet.",
            "execution is optional; omit it for FULL_XSSF / SUMMARY / DO_NOT_CALCULATE,"
                + " or supply only the nested execution.mode, execution.journal, and"
                + " execution.calculation fields that need non-default behavior."));
  }

  static HelpSurfaceReport.CliDefinitionSection limits() {
    return new HelpSurfaceReport.CliDefinitionSection(
        "Limits",
        List.of(
            new HelpSurfaceReport.DefinitionEntry(
                "File format", ".xlsx only; .xls, .xlsm, and .xlsb are rejected."),
            new HelpSurfaceReport.DefinitionEntry(
                "Sheet names",
                "1 to 31 characters; reject : \\ / ? * [ ] and leading/trailing" + " apostrophes."),
            new HelpSurfaceReport.DefinitionEntry(
                "GET_CELLS addresses",
                "addresses must not exceed 250,000; the cap keys off the exact returned cell"
                    + " count."),
            new HelpSurfaceReport.DefinitionEntry(
                "GET_WINDOW cell count",
                "rowCount * columnCount must not exceed 250,000; sparse output lowers"
                    + " blank-heavy payloads but the request cap still keys off the full"
                    + " rectangle."),
            new HelpSurfaceReport.DefinitionEntry(
                "GET_SHEET_SCHEMA cells",
                "rowCount * columnCount must not exceed 250,000 because schema inference"
                    + " still examines the full rectangular sample."),
            new HelpSurfaceReport.DefinitionEntry(
                "Worksheet materialization",
                "cell-range and row-materializing actions together must not materialize more"
                    + " than 250,000 worksheet items per plan."),
            new HelpSurfaceReport.DefinitionEntry(
                "Request JSON size",
                "request JSON must not exceed 16 MiB ("
                    + GridGrindRequestSurfaceContractText.requestDocumentLimitBytes()
                    + " bytes); authored text and binary payloads may instead use"
                    + " UTF8_FILE, FILE, or STANDARD_INPUT sources."),
            new HelpSurfaceReport.DefinitionEntry(
                "EVENT_READ mode", BackendCapabilities.eventRead().catalogSummary()),
            new HelpSurfaceReport.DefinitionEntry(
                "STREAMING_WRITE mode", BackendCapabilities.streamingWrite().catalogSummary()),
            new HelpSurfaceReport.DefinitionEntry(
                "OOXML write encryption", GridGrindOoxmlWriteEncryptionContractText.limitSummary()),
            new HelpSurfaceReport.DefinitionEntry(
                "Column widthCharacters", "authored widthCharacters > 0 and <= 255 (Excel limit)."),
            new HelpSurfaceReport.DefinitionEntry(
                "Default sheet sizing",
                "authored defaultColumnWidth must be > 0 and <= 255;"
                    + " authored defaultRowHeightPoints must be > 0 and <= "
                    + dev.erst.gridgrind.core.vocabulary.ExcelSheetLayoutLimits
                        .MAX_ROW_HEIGHT_POINTS
                    + " (Excel limits)."),
            new HelpSurfaceReport.DefinitionEntry(
                "Row heightPoints",
                "authored heightPoints > 0 and <= "
                    + dev.erst.gridgrind.core.vocabulary.ExcelSheetLayoutLimits
                        .MAX_ROW_HEIGHT_POINTS
                    + " (Excel row height limit)."),
            new HelpSurfaceReport.DefinitionEntry(
                "Zoom percent",
                dev.erst.gridgrind.core.vocabulary.ExcelSheetLayoutLimits.MIN_ZOOM_PERCENT
                    + " to "
                    + dev.erst.gridgrind.core.vocabulary.ExcelSheetLayoutLimits.MAX_ZOOM_PERCENT
                    + " inclusive (Excel zoom limit)."),
            new HelpSurfaceReport.DefinitionEntry(
                "Row structural edits",
                "rejected when they would move tables, sheet autofilters, or data"
                    + " validations; deletes/shifts also reject destructive range-backed"
                    + " named ranges."),
            new HelpSurfaceReport.DefinitionEntry(
                "Column structural edits",
                "rejected when they would move tables, sheet autofilters, or data"
                    + " validations; deletes/shifts also reject destructive range-backed"
                    + " named ranges; all column edits reject any workbook formulas or"
                    + " formula-defined names."),
            new HelpSurfaceReport.DefinitionEntry(
                "Chart mutations",
                "SET_CHART authors AREA, AREA_3D, BAR, BAR_3D, DOUGHNUT, LINE,"
                    + " LINE_3D, PIE, PIE_3D, RADAR, SCATTER, SURFACE, and SURFACE_3D;"
                    + " unsupported loaded chart detail is preserved on unrelated edits"
                    + " and rejected for authoritative mutation."),
            new HelpSurfaceReport.DefinitionEntry(
                "Chart title formulas",
                "SET_CHART title FORMULA and series.title FORMULA must resolve to one"
                    + " cell, directly or through one defined name."),
            new HelpSurfaceReport.DefinitionEntry(
                "Array formulas",
                "SET_ARRAY_FORMULA authors one contiguous single-cell or multi-cell array"
                    + " formula group; CLEAR_ARRAY_FORMULA may target any member cell and"
                    + " removes the whole stored group."),
            new HelpSurfaceReport.DefinitionEntry(
                "Drawing validation",
                "failed SET_SHAPE / SET_CHART validation leaves existing drawing state"
                    + " unchanged and creates no partial artifacts."),
            new HelpSurfaceReport.DefinitionEntry(
                "DATE / DATE_TIME inputs",
                "stored as numeric serial; default VALUE readback keeps type=NUMBER, and the"
                    + " TEMPORAL facet reveals DATE / TIME / DATE_TIME semantics when that"
                    + " distinction matters."),
            new HelpSurfaceReport.DefinitionEntry(
                "Formula authoring",
                GridGrindInspectionContractText.formulaAuthoringLimitSummary()),
            new HelpSurfaceReport.DefinitionEntry(
                "Loaded formula support",
                GridGrindInspectionContractText.loadedFormulaSupportSummary())));
  }
}
