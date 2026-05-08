package dev.erst.gridgrind.cli.discovery;

import static dev.erst.gridgrind.cli.discovery.GridGrindTaskEntrySupport.phase;
import static dev.erst.gridgrind.cli.discovery.GridGrindTaskEntrySupport.profile;
import static dev.erst.gridgrind.cli.discovery.GridGrindTaskEntrySupport.ref;
import static dev.erst.gridgrind.cli.discovery.GridGrindTaskEntrySupport.signals;
import static dev.erst.gridgrind.cli.discovery.GridGrindTaskEntrySupport.task;
import static dev.erst.gridgrind.cli.discovery.GridGrindTaskEntrySupport.workflow;

import java.util.List;

/** CLI-owned task descriptors for report, dashboard, intake, and pivot workflows. */
final class GridGrindReportingTaskDefinitions {
  private GridGrindReportingTaskDefinitions() {}

  static List<TaskEntry> entries() {
    return List.of(tabularReport(), dashboard(), dataEntryWorkflow(), pivotReport());
  }

  private static TaskEntry tabularReport() {
    return task(
        "TABULAR_REPORT",
        "Create a structured worksheet report with typed cells, table semantics, and factual"
            + " readback.",
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
        List.of("office", "reporting", "table", "report", "save"),
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
            "Assertions on key balances or counts."),
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

  private static TaskEntry dashboard() {
    return task(
        "DASHBOARD",
        "Assemble an executive dashboard from reusable named surfaces and supported charts.",
        profile(
            TaskSourceMode.NEW_WORKBOOK,
            TaskPersistenceMode.SAVE_AS,
            TaskMutationMode.MUTATING,
            TaskAssetMode.SELF_CONTAINED),
        signals(
            List.of(
                TaskInputKind.TARGET_SHEET_NAMES,
                TaskInputKind.TARGET_OBJECT_NAMES,
                TaskInputKind.CELL_OR_RANGE_COORDINATES,
                TaskInputKind.PERSISTENCE_TARGET_PATH),
            List.of(
                TaskVerificationKind.FACT_READBACK,
                TaskVerificationKind.ASSERTION_CHECKS,
                TaskVerificationKind.HEALTH_ANALYSIS)),
        List.of("office", "dashboard", "charts", "summary", "kpi"),
        List.of(
            "Summary sheets and KPI surfaces are intentionally structured.",
            "Reusable named surfaces back formulas or charts instead of fragile copied ranges.",
            "Charts are authored and then read back through the same contract surface."),
        List.of(
            "Metric definitions and source ranges.",
            "Chart names and anchors.",
            "Target persistence path when the dashboard must be saved."),
        List.of(
            "Assertions that required dashboard entities exist.",
            "Workbook-health analysis after authoring.",
            "Named-range-backed chart series for reusable models."),
        workflow(
            List.of(
                phase(
                    TaskPhasePurpose.PREPARE,
                    "Assemble Summary Sheets",
                    "Create the dashboard canvas and key text or formula cells first.",
                    List.of(
                        ref("sourceTypes", "NEW"),
                        ref("mutationActionTypes", "ENSURE_SHEET"),
                        ref("mutationActionTypes", "SET_CELL"),
                        ref("mutationActionTypes", "SET_RANGE")),
                    List.of("Keep summary layout intentional so later chart anchors are stable.")),
                phase(
                    TaskPhasePurpose.AUTHOR,
                    "Define Reusable Model Surfaces",
                    "Create named ranges that charts and formulas can depend on.",
                    List.of(
                        ref("mutationActionTypes", "SET_NAMED_RANGE"),
                        ref("mutationActionTypes", "SET_CHART")),
                    List.of("Named surfaces reduce accidental drift when the dashboard evolves.")),
                phase(
                    TaskPhasePurpose.VERIFY,
                    "Author And Inspect Visuals",
                    "Create supported charts and verify that the expected entities exist.",
                    List.of(
                        ref("inspectionQueryTypes", "GET_CHARTS"),
                        ref("assertionTypes", "EXPECT_CHART_PRESENT"),
                        ref("persistenceTypes", "SAVE_AS")),
                    List.of(
                        "Use factual chart readback to confirm what the workbook now contains."))),
            List.of(
                "SET_CHART supports the authoritative simple-chart family listed in the"
                    + " protocol catalog.",
                "Chart title and series FORMULA titles must resolve to one cell.",
                "Unsupported loaded chart detail is preserved on unrelated edits but is not"
                    + " available for authoritative mutation.")));
  }

  private static TaskEntry dataEntryWorkflow() {
    return task(
        "DATA_ENTRY_WORKFLOW",
        "Build one repeatable intake worksheet for row-oriented data entry with validations,"
            + " comments, and later factual inspection.",
        profile(
            TaskSourceMode.NEW_WORKBOOK,
            TaskPersistenceMode.SAVE_AS,
            TaskMutationMode.MUTATING,
            TaskAssetMode.SELF_CONTAINED),
        signals(
            List.of(
                TaskInputKind.TARGET_SHEET_NAMES,
                TaskInputKind.VALIDATION_RULES,
                TaskInputKind.PERSISTENCE_TARGET_PATH),
            List.of(TaskVerificationKind.FACT_READBACK, TaskVerificationKind.ASSERTION_CHECKS)),
        List.of("office", "data entry", "validation", "intake", "worksheet"),
        List.of(
            "Operators get one workbook surface designed for repeated entry instead of ad hoc"
                + " edits.",
            "Validation rules, comments, and protection settings are part of the authored"
                + " workflow.",
            "The result can be inspected or asserted after authoring."),
        List.of(
            "Target sheet structure and protected cells.",
            "Validation ranges and allowed values.",
            "Save target when the intake workbook must be persisted."),
        List.of(
            "Sheet protection after authoring.",
            "Comments or prompts that explain the allowed entry flow.",
            "Assertions on the protected or validated workbook surface."),
        workflow(
            List.of(
                phase(
                    TaskPhasePurpose.PREPARE,
                    "Prepare The Intake Sheet",
                    "Create the sheet skeleton, labels, and writable cells first.",
                    List.of(
                        ref("sourceTypes", "NEW"),
                        ref("mutationActionTypes", "ENSURE_SHEET"),
                        ref("mutationActionTypes", "SET_RANGE"),
                        ref("mutationActionTypes", "SET_CELL")),
                    List.of(
                        "Start with the worksheet shape before layering validations or"
                            + " protection.")),
                phase(
                    TaskPhasePurpose.AUTHOR,
                    "Author Guardrails",
                    "Add validations, comments, prompts, and protection that shape entry"
                        + " behavior.",
                    List.of(
                        ref("mutationActionTypes", "SET_DATA_VALIDATION"),
                        ref("mutationActionTypes", "SET_COMMENT"),
                        ref("mutationActionTypes", "SET_WORKBOOK_PROTECTION"),
                        ref("persistenceTypes", "SAVE_AS")),
                    List.of("Validation and protection belong to the authored model, not a memo.")),
                phase(
                    TaskPhasePurpose.VERIFY,
                    "Inspect The Surface",
                    "Read back validations, comments, and summary facts before shipping the"
                        + " workbook.",
                    List.of(
                        ref("inspectionQueryTypes", "GET_DATA_VALIDATIONS"),
                        ref("inspectionQueryTypes", "GET_COMMENTS"),
                        ref("inspectionQueryTypes", "GET_WORKBOOK_SUMMARY")),
                    List.of("Factual rereads catch drift before the workbook reaches operators."))),
            List.of(
                "Protection settings can restrict later mutation flows unless you plan for"
                    + " them.",
                "Validation formulas and lists must fit the supported POI-backed contract"
                    + " shape.",
                "Large instructional text belongs in external text sources instead of huge"
                    + " inline literals.")));
  }

  private static TaskEntry pivotReport() {
    return task(
        "PIVOT_REPORT",
        "Create one pivot-ready report flow from source data through named tables and pivot-table"
            + " inspection.",
        profile(
            TaskSourceMode.NEW_WORKBOOK,
            TaskPersistenceMode.SAVE_AS,
            TaskMutationMode.MUTATING,
            TaskAssetMode.SELF_CONTAINED),
        signals(
            List.of(
                TaskInputKind.TARGET_SHEET_NAMES,
                TaskInputKind.TARGET_OBJECT_NAMES,
                TaskInputKind.TABULAR_SOURCE_ROWS,
                TaskInputKind.PERSISTENCE_TARGET_PATH),
            List.of(
                TaskVerificationKind.FACT_READBACK,
                TaskVerificationKind.ASSERTION_CHECKS,
                TaskVerificationKind.HEALTH_ANALYSIS)),
        List.of("office", "pivot", "summary", "analysis", "table"),
        List.of(
            "Source data is modeled as one table before pivot logic depends on it.",
            "Pivot tables are authored by name instead of fragile workbook surgery.",
            "Readback confirms the resulting pivot structure."),
        List.of(
            "Source rows and summary dimensions.",
            "Pivot destination sheet and placement.",
            "Persistence target when the workbook must be saved."),
        List.of(
            "Workbook findings after authoring.",
            "Charting or dashboard follow-ups built from the pivot output.",
            "Assertions that the expected pivot table exists."),
        workflow(
            List.of(
                phase(
                    TaskPhasePurpose.PREPARE,
                    "Model Source Data",
                    "Create one table-backed source surface for later summarization.",
                    List.of(
                        ref("sourceTypes", "NEW"),
                        ref("mutationActionTypes", "ENSURE_SHEET"),
                        ref("mutationActionTypes", "SET_RANGE"),
                        ref("mutationActionTypes", "SET_TABLE")),
                    List.of("Pivot work starts from explicit tabular semantics.")),
                phase(
                    TaskPhasePurpose.AUTHOR,
                    "Author The Pivot",
                    "Create the pivot table and place it on a durable destination sheet.",
                    List.of(
                        ref("mutationActionTypes", "SET_PIVOT_TABLE"),
                        ref("mutationActionTypes", "ENSURE_SHEET"),
                        ref("persistenceTypes", "SAVE_AS")),
                    List.of("Author the destination shape before adding later visual layers.")),
                phase(
                    TaskPhasePurpose.VERIFY,
                    "Inspect And Verify",
                    "Read back pivot facts and broader workbook health after authoring.",
                    List.of(
                        ref("inspectionQueryTypes", "GET_PIVOT_TABLES"),
                        ref("inspectionQueryTypes", "ANALYZE_WORKBOOK_FINDINGS"),
                        ref("assertionTypes", "EXPECT_PIVOT_TABLE_PRESENT")),
                    List.of("Use pivot readback to confirm the authored summary surface."))),
            List.of(
                "Pivot authoring depends on stable source headers and range shapes.",
                "Calculated fields and unsupported advanced pivot features may require later"
                    + " manual workbook refinement.",
                "If the workflow grows visual layers, keep the pivot and chart"
                    + " responsibilities separate.")));
  }
}
