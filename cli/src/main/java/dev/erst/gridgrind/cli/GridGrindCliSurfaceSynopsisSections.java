package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.contract.catalog.GridGrindExecutionModeMetadata;
import dev.erst.gridgrind.contract.catalog.GridGrindInspectionContractText;
import dev.erst.gridgrind.contract.catalog.GridGrindOoxmlWriteEncryptionContractText;
import dev.erst.gridgrind.contract.catalog.GridGrindRequestSurfaceContractText;
import java.util.List;

/** Synopsis-first CLI help sections. */
final class GridGrindCliSurfaceSynopsisSections {
  private GridGrindCliSurfaceSynopsisSections() {}

  static CliSurface.CliSection usage() {
    return new CliSurface.CliSection(
        "Usage",
        List.of(
            "gridgrind --request <path> [--temp-root <path>] [--response <path>] [--pretty]",
            "gridgrind --execution-root <path> [--temp-root <path>] [--response <path>]"
                + " [--pretty]"
                + " < request.json",
            "gridgrind --doctor-request --request <path> [--temp-root <path>]"
                + " [--response <path>] [--pretty]",
            "gridgrind --doctor-request --execution-root <path> [--temp-root <path>]"
                + " [--response <path>] [--pretty] < request.json",
            "gridgrind --print-request-template [--response <path>] [--pretty]",
            "gridgrind --print-recipe --lookup <id> [--response <path>] [--pretty]",
            "gridgrind --print-recipe-catalog [--lookup <id>] [--response <path>] [--pretty]",
            "gridgrind --print-recipe-keyword-match --query <text> [--response <path>]"
                + " [--pretty]",
            "gridgrind --print-protocol-catalog [--response <path>] [--pretty]",
            "gridgrind --print-protocol-catalog --lookup <lookup-id> [--response"
                + " <path>] [--pretty]",
            "gridgrind --print-protocol-catalog --search <text> [--response <path>]"
                + " [--pretty]",
            "gridgrind --help | -h | help [--response <path>] [--format <text|structured>]"
                + " [--pretty]",
            "gridgrind --help-protocol | help-protocol [--response <path>] [--format"
                + " <text|structured>] [--pretty]",
            "gridgrind --help-guidance | help-guidance [--response <path>] [--format"
                + " <text|structured>] [--pretty]",
            "gridgrind --version | version [--response <path>] [--format <text|structured>]"
                + " [--pretty]",
            "gridgrind --license | license [--response <path>] [--format <text|structured>]"
                + " [--pretty]"));
  }

  static CliSurface.CliWorkflowSection workflows() {
    return new CliSurface.CliWorkflowSection(
        "Workflows",
        List.of(
            new CliSurface.WorkflowEntry(
                "Discover What To Send",
                List.of(
                    "List every built-in recipe: gridgrind --print-recipe-catalog"
                        + " --response recipes.json",
                    "Get one executable starter scenario: gridgrind --print-recipe"
                        + " --lookup DASHBOARD --response dashboard-request.json",
                    "Get the compact protocol-catalog index: gridgrind"
                        + " --print-protocol-catalog --response protocol-index.json",
                    "Search exact protocol shapes: gridgrind --print-protocol-catalog"
                        + " --search \"chart title\" --response catalog-search.json",
                    "Resolve one scoped catalog group: gridgrind --print-protocol-catalog"
                        + " --lookup calculationStrategyTypes --response"
                        + " calculation-strategy-types.json")),
            new CliSurface.WorkflowEntry(
                "Draft And Preflight",
                List.of(
                    "Start from the minimal request: gridgrind --print-request-template"
                        + " --response request.json",
                    "For stdin-driven execution or doctoring, pass one explicit"
                        + " --execution-root so request-owned paths resolve from one"
                        + " explicit invocation directory.",
                    "Use --print-recipe --lookup <id> when you want one shipped"
                        + " executable example or task-starter scenario instead of"
                        + " building from scratch.",
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
                        + " together with one explicit SAVE_AS.ifExists policy,"
                        + " and persistence.type=NONE when the run is read-only or"
                        + " diagnostic."))));
  }

  static CliSurface.CliSection execution() {
    return new CliSurface.CliSection(
        "Execution",
        List.of(
            "GridGrind executes ordered steps in sequence, then saves the workbook (unless"
                + " persistence is NONE); if any step fails, no workbook is written.",
            "source.type=NEW starts with zero sheets; ENSURE_SHEET creates the first" + " sheet.",
            "execution is optional; omit it for FULL_XSSF / SUMMARY / DO_NOT_CALCULATE,"
                + " or supply only the nested execution.mode, execution.journal, and"
                + " execution.calculation fields that need non-default behavior."));
  }

  static CliSurface.CliDefinitionSection limits() {
    return new CliSurface.CliDefinitionSection(
        "Limits",
        List.of(
            new CliSurface.DefinitionEntry(
                "File format", ".xlsx only; .xls, .xlsm, and .xlsb are rejected."),
            new CliSurface.DefinitionEntry(
                "Sheet names",
                "1 to 31 characters; reject : \\ / ? * [ ] and leading/trailing" + " apostrophes."),
            new CliSurface.DefinitionEntry(
                "GET_CELLS addresses",
                "addresses must not exceed 250,000; the cap keys off the exact returned cell"
                    + " count."),
            new CliSurface.DefinitionEntry(
                "GET_WINDOW cell count",
                "rowCount * columnCount must not exceed 250,000; sparse output lowers"
                    + " blank-heavy payloads but the request cap still keys off the full"
                    + " rectangle."),
            new CliSurface.DefinitionEntry(
                "GET_SHEET_SCHEMA cells",
                "rowCount * columnCount must not exceed 250,000 because schema inference"
                    + " still examines the full rectangular sample."),
            new CliSurface.DefinitionEntry(
                "Request JSON size",
                "request JSON must not exceed 16 MiB ("
                    + GridGrindRequestSurfaceContractText.requestDocumentLimitBytes()
                    + " bytes); authored text and binary payloads may instead use"
                    + " UTF8_FILE, FILE, or STANDARD_INPUT sources."),
            new CliSurface.DefinitionEntry(
                "EVENT_READ mode", GridGrindExecutionModeMetadata.eventRead().catalogSummary()),
            new CliSurface.DefinitionEntry(
                "STREAMING_WRITE mode",
                GridGrindExecutionModeMetadata.streamingWrite().catalogSummary()),
            new CliSurface.DefinitionEntry(
                "OOXML write encryption", GridGrindOoxmlWriteEncryptionContractText.limitSummary()),
            new CliSurface.DefinitionEntry(
                "Column widthCharacters", "authored widthCharacters > 0 and <= 255 (Excel limit)."),
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
                    + dev.erst.gridgrind.excel.foundation.ExcelSheetLayoutLimits.MAX_ZOOM_PERCENT
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
                    + " and rejected for authoritative mutation."),
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
                "stored as numeric serial; default VALUE readback keeps type=NUMBER, and the"
                    + " TEMPORAL facet reveals DATE / TIME / DATE_TIME semantics when that"
                    + " distinction matters."),
            new CliSurface.DefinitionEntry(
                "Formula authoring",
                GridGrindInspectionContractText.formulaAuthoringLimitSummary()),
            new CliSurface.DefinitionEntry(
                "Loaded formula support",
                GridGrindInspectionContractText.loadedFormulaSupportSummary())));
  }
}
