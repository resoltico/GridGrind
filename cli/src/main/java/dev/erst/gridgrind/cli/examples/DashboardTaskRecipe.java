package dev.erst.gridgrind.cli.examples;

import static dev.erst.gridgrind.cli.examples.GridGrindTaskRecipeSupport.discovery;
import static dev.erst.gridgrind.cli.examples.GridGrindTaskRecipeSupport.intent;
import static dev.erst.gridgrind.cli.examples.GridGrindTaskRecipeSupport.narrative;
import static dev.erst.gridgrind.cli.examples.GridGrindTaskRecipeSupport.phase;
import static dev.erst.gridgrind.cli.examples.GridGrindTaskRecipeSupport.profile;
import static dev.erst.gridgrind.cli.examples.GridGrindTaskRecipeSupport.ref;
import static dev.erst.gridgrind.cli.examples.GridGrindTaskRecipeSupport.selfContainedStarter;
import static dev.erst.gridgrind.cli.examples.GridGrindTaskRecipeSupport.signals;
import static dev.erst.gridgrind.cli.examples.GridGrindTaskRecipeSupport.task;
import static dev.erst.gridgrind.cli.examples.GridGrindTaskRecipeSupport.workflow;

import dev.erst.gridgrind.cli.discovery.TaskArtifactKind;
import dev.erst.gridgrind.cli.discovery.TaskAssetMode;
import dev.erst.gridgrind.cli.discovery.TaskGoalKind;
import dev.erst.gridgrind.cli.discovery.TaskInputKind;
import dev.erst.gridgrind.cli.discovery.TaskMutationMode;
import dev.erst.gridgrind.cli.discovery.TaskPersistenceMode;
import dev.erst.gridgrind.cli.discovery.TaskPhasePurpose;
import dev.erst.gridgrind.cli.discovery.TaskSourceMode;
import dev.erst.gridgrind.cli.discovery.TaskVerificationKind;
import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.DrawingMutationAction;
import dev.erst.gridgrind.contract.action.StructuredMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.WorkbookAssetIntrospectionQuery;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import java.util.List;

/** Published task recipe for executive dashboard workflows. */
final class DashboardTaskRecipe {
  private DashboardTaskRecipe() {}

  static GridGrindTaskRecipeDefinition definition() {
    return GridGrindTaskRecipeSupport.definition(
        task(
            "DASHBOARD",
            selfContainedStarter("DASHBOARD"),
            List.of("office", "dashboard", "charts", "summary", "kpi", "sales", "revenue"),
            discovery(
                List.of(
                    "dashboard",
                    "charts",
                    "chart",
                    "kpi",
                    "scorecard",
                    "executive summary",
                    "sales",
                    "revenue"),
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
                        List.of(
                            "Keep summary layout intentional so later chart anchors are stable.")),
                    phase(
                        TaskPhasePurpose.AUTHOR,
                        "Define Reusable Model Surfaces",
                        "Create named ranges that charts and formulas can depend on.",
                        List.of(
                            ref("mutationActionTypes", "SET_NAMED_RANGE"),
                            ref("mutationActionTypes", "SET_CHART")),
                        List.of(
                            "Named surfaces reduce accidental drift when the dashboard evolves.")),
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
                    "SET_CHART supports the authoritative simple-chart family listed in the protocol catalog.",
                    "Chart title and series FORMULA titles must resolve to one cell.",
                    "Unsupported loaded chart detail is preserved on unrelated edits but is not available for authoritative mutation."))),
        starterPlan());
  }

  private static WorkbookPlan starterPlan() {
    String taskId = "DASHBOARD";
    return ExampleWorkbookPlans.defaultExecutionPlan(
        TaskStarterRecipeSupport.taskPlanId(taskId),
        new WorkbookPlan.WorkbookSource.New(),
        ExampleWorkbookPlans.saveAs(TaskStarterRecipeSupport.taskWorkbookPath(taskId)),
        ExampleSteps.step(
            "ensure-ops", ExampleSelectors.sheet("Ops"), new WorkbookMutationAction.EnsureSheet()),
        ExampleSteps.step(
            "seed-ops-range",
            ExampleSelectors.range("Ops", "A1:C4"),
            new CellMutationAction.SetRange(
                ExampleCellValues.rows(
                    ExampleCellValues.row(
                        ExampleCellValues.text("Month"),
                        ExampleCellValues.text("Plan"),
                        ExampleCellValues.text("Actual")),
                    ExampleCellValues.row(
                        ExampleCellValues.text("Jan"),
                        ExampleCellValues.number(10.0d),
                        ExampleCellValues.number(12.0d)),
                    ExampleCellValues.row(
                        ExampleCellValues.text("Feb"),
                        ExampleCellValues.number(18.0d),
                        ExampleCellValues.number(17.0d)),
                    ExampleCellValues.row(
                        ExampleCellValues.text("Mar"),
                        ExampleCellValues.number(15.0d),
                        ExampleCellValues.number(16.0d))))),
        ExampleSteps.step(
            "define-chart-categories",
            new NamedRangeSelector.WorkbookScope("DashboardCategories"),
            new StructuredMutationAction.SetNamedRange(
                "DashboardCategories",
                new dev.erst.gridgrind.contract.dto.NamedRangeScope.Workbook(),
                dev.erst.gridgrind.contract.dto.NamedRangeTarget.range("Ops", "A2:A4"))),
        ExampleSteps.step(
            "define-dashboard-actual",
            new NamedRangeSelector.WorkbookScope("DashboardActual"),
            new StructuredMutationAction.SetNamedRange(
                "DashboardActual",
                new dev.erst.gridgrind.contract.dto.NamedRangeScope.Workbook(),
                dev.erst.gridgrind.contract.dto.NamedRangeTarget.range("Ops", "C2:C4"))),
        ExampleSteps.step(
            "author-dashboard-chart",
            ExampleSelectors.sheet("Ops"),
            new DrawingMutationAction.SetChart(
                ExampleChartInputs.clusteredColumnComparisonChart(
                    "DashboardChart",
                    ExampleDrawingAnchors.anchor(4, 0, 8, 12),
                    "Operations Dashboard",
                    "DashboardCategories",
                    "Plan",
                    "Ops!$B$2:$B$4",
                    "Actual",
                    "DashboardActual"))),
        ExampleSteps.read(
            "read-dashboard-charts",
            new dev.erst.gridgrind.contract.selector.ChartSelector.AllOnSheet("Ops"),
            new WorkbookAssetIntrospectionQuery.GetCharts()),
        ExampleSteps.assertStep(
            "assert-dashboard-chart",
            new dev.erst.gridgrind.contract.selector.ChartSelector.ByName("Ops", "DashboardChart"),
            new dev.erst.gridgrind.contract.assertion.PresenceAssertion.ChartPresent()));
  }
}
