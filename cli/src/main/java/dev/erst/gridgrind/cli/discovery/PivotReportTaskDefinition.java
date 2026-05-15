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

/** Task descriptor for pivot-table reporting workflows. */
final class PivotReportTaskDefinition {
  private PivotReportTaskDefinition() {}

  static TaskEntry entry() {
    return task(
        "PIVOT_REPORT",
        discovery(
            List.of("pivot", "pivot table", "summary", "analysis", "cross-tab"),
            List.of("office", "pivot", "summary", "analysis", "table"),
            intent(
                List.of(TaskGoalKind.AUTHOR, TaskGoalKind.VERIFY, TaskGoalKind.ANALYZE),
                List.of(
                    TaskArtifactKind.WORKBOOK,
                    TaskArtifactKind.TABLE,
                    TaskArtifactKind.PIVOT_TABLE))),
        narrative(
            "Create one pivot-ready report flow from source data through named tables and"
                + " pivot-table inspection.",
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
                "Assertions that the expected pivot table exists.")),
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
