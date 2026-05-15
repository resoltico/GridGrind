package dev.erst.gridgrind.cli.discovery;

import static dev.erst.gridgrind.cli.discovery.GridGrindTaskEntrySupport.discovery;
import static dev.erst.gridgrind.cli.discovery.GridGrindTaskEntrySupport.intent;
import static dev.erst.gridgrind.cli.discovery.GridGrindTaskEntrySupport.narrative;
import static dev.erst.gridgrind.cli.discovery.GridGrindTaskEntrySupport.phase;
import static dev.erst.gridgrind.cli.discovery.GridGrindTaskEntrySupport.profile;
import static dev.erst.gridgrind.cli.discovery.GridGrindTaskEntrySupport.ref;
import static dev.erst.gridgrind.cli.discovery.GridGrindTaskEntrySupport.signals;
import static dev.erst.gridgrind.cli.discovery.GridGrindTaskEntrySupport.task;
import static dev.erst.gridgrind.cli.discovery.GridGrindTaskEntrySupport.workflow;

import java.util.List;

/** Task descriptor for structured worksheet reporting workflows. */
final class TabularReportTaskDefinition {
  private TabularReportTaskDefinition() {}

  static TaskEntry entry() {
    return task(
        "TABULAR_REPORT",
        discovery(
            List.of("report", "spreadsheet", "worksheet", "budget", "tracker", "ledger", "table"),
            List.of("office", "reporting", "table", "report", "save"),
            intent(
                List.of(TaskGoalKind.AUTHOR, TaskGoalKind.VERIFY),
                List.of(
                    TaskArtifactKind.WORKBOOK,
                    TaskArtifactKind.SHEET,
                    TaskArtifactKind.CELL,
                    TaskArtifactKind.TABLE))),
        narrative(
            "Create a structured worksheet report with typed cells, table semantics, and factual"
                + " readback.",
            List.of(
                "Sheet structure is created intentionally instead of ad hoc cell drift.",
                "Rows can be modeled as a table so filtering and later readback stay authoritative.",
                "Critical facts can be inspected before the workbook is persisted."),
            List.of(
                "Sheet name and header layout.",
                "Typed row values or source-backed payload files for larger authored content.",
                "Persistence target when the result must be saved."),
            List.of(
                "Totals rows and formula-backed summaries.",
                "Cell styling and print-layout refinements.",
                "Assertions on key balances or counts.")),
        profile(
            TaskSourceMode.NEW_WORKBOOK,
            TaskPersistenceMode.SAVE_AS,
            TaskMutationMode.MUTATING,
            TaskAssetMode.SELF_CONTAINED),
        signals(
            List.of(
                TaskInputKind.TARGET_SHEET_NAMES,
                TaskInputKind.TABULAR_SOURCE_ROWS,
                TaskInputKind.PERSISTENCE_TARGET_PATH),
            List.of(TaskVerificationKind.FACT_READBACK, TaskVerificationKind.ASSERTION_CHECKS)),
        workflow(
            List.of(
                phase(
                    TaskPhasePurpose.PREPARE,
                    "Lay Out The Workbook",
                    "Create sheets, headers, and fixed structure before data rows arrive.",
                    List.of(
                        ref("sourceTypes", "NEW"),
                        ref("mutationActionTypes", "ENSURE_SHEET"),
                        ref("mutationActionTypes", "SET_RANGE")),
                    List.of(
                        "Use one intentional sheet skeleton instead of scattered cell writes.")),
                phase(
                    TaskPhasePurpose.AUTHOR,
                    "Model The Table",
                    "Move the sheet from loose cells into tabular semantics.",
                    List.of(
                        ref("mutationActionTypes", "SET_TABLE"),
                        ref("mutationActionTypes", "AUTO_SIZE_COLUMNS"),
                        ref("persistenceTypes", "SAVE_AS")),
                    List.of("Prefer one table per logical data region.")),
                phase(
                    TaskPhasePurpose.VERIFY,
                    "Verify And Inspect",
                    "Read back the cells or workbook facts that make the report trustworthy.",
                    List.of(
                        ref("inspectionQueryTypes", "GET_CELLS"),
                        ref("inspectionQueryTypes", "GET_TABLES"),
                        ref("inspectionQueryTypes", "GET_WORKBOOK_SUMMARY")),
                    List.of("Use factual readback instead of assuming writes landed correctly."))),
            List.of(
                "Large authored literals belong in UTF8_FILE, FILE, or STANDARD_INPUT sources"
                    + " instead of huge inline JSON.",
                "Table headers must remain nonblank and unique.",
                "Formula authoring is scalar-only unless you use SET_ARRAY_FORMULA.")));
  }
}
