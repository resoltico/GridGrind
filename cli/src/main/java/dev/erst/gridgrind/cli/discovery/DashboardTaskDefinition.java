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

/** Task descriptor for executive dashboard workflows. */
final class DashboardTaskDefinition {
  private DashboardTaskDefinition() {}

  static TaskEntry entry() {
    return task(
        "DASHBOARD",
        discovery(
            List.of("dashboard", "charts", "chart", "kpi", "scorecard", "executive summary"),
            List.of("office", "dashboard", "charts", "summary", "kpi"),
            intent(
                List.of(TaskGoalKind.AUTHOR, TaskGoalKind.VERIFY, TaskGoalKind.ANALYZE),
                List.of(
                    TaskArtifactKind.WORKBOOK,
                    TaskArtifactKind.SHEET,
                    TaskArtifactKind.CHART,
                    TaskArtifactKind.NAMED_RANGE))),
        narrative(
            "Assemble an executive dashboard from reusable named surfaces and supported charts.",
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
                "Named-range-backed chart series for reusable models.")),
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
}
