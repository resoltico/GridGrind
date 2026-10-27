package dev.erst.gridgrind.core.discovery;

import dev.erst.gridgrind.core.model.table.TableSpec;
import dev.erst.gridgrind.core.model.table.TableStyleSpec;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.InspectionAnalysisQuery;
import dev.erst.gridgrind.core.plan.PivotTableSelector;
import dev.erst.gridgrind.core.plan.PresenceAssertion;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookAssetIntrospectionQuery;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.core.vocabulary.TableName;
import java.util.List;

/** Published task recipe for pivot-table reporting workflows. */
final class PivotReportTaskRecipe {
  private PivotReportTaskRecipe() {}

  static TaskRecipeTemplate definition() {
    return TaskRecipes.definition(
        TaskRecipes.task(
            "PIVOT_REPORT",
            TaskRecipes.selfContainedStarter("PIVOT_REPORT"),
            List.of("office", "pivot", "summary", "analysis", "table"),
            TaskRecipes.discovery(
                List.of("pivot", "pivot table", "summary", "analysis", "cross-tab"),
                TaskRecipes.intent(
                    List.of(TaskGoalKind.AUTHOR, TaskGoalKind.VERIFY, TaskGoalKind.ANALYZE),
                    List.of(
                        TaskArtifactKind.WORKBOOK,
                        TaskArtifactKind.TABLE,
                        TaskArtifactKind.PIVOT_TABLE))),
            TaskRecipes.narrative(
                "Create one pivot-ready report flow from source data through named tables and pivot-table inspection.",
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
            TaskRecipes.profile(
                TaskSourceMode.NEW_WORKBOOK,
                TaskPersistenceMode.SAVE_AS,
                TaskMutationMode.MUTATING,
                TaskAssetMode.SELF_CONTAINED),
            TaskRecipes.signals(
                List.of(
                    TaskInputKind.TARGET_SHEET_NAMES,
                    TaskInputKind.TARGET_OBJECT_NAMES,
                    TaskInputKind.TABULAR_SOURCE_ROWS,
                    TaskInputKind.PERSISTENCE_TARGET_PATH),
                List.of(
                    TaskVerificationKind.FACT_READBACK,
                    TaskVerificationKind.ASSERTION_CHECKS,
                    TaskVerificationKind.HEALTH_ANALYSIS)),
            TaskRecipes.workflow(
                List.of(
                    TaskRecipes.phase(
                        TaskPhasePurpose.PREPARE,
                        "Model Source Data",
                        "Create one table-backed source surface for later summarization.",
                        List.of(
                            TaskRecipes.ref("workbookSourceTypes", "NEW"),
                            TaskRecipes.ref("mutationActionTypes", "ENSURE_SHEET"),
                            TaskRecipes.ref("mutationActionTypes", "SET_RANGE"),
                            TaskRecipes.ref("mutationActionTypes", "SET_TABLE")),
                        List.of("Pivot work starts from explicit tabular semantics.")),
                    TaskRecipes.phase(
                        TaskPhasePurpose.AUTHOR,
                        "Author The Pivot",
                        "Create the pivot table and place it on a durable destination sheet.",
                        List.of(
                            TaskRecipes.ref("mutationActionTypes", "SET_PIVOT_TABLE"),
                            TaskRecipes.ref("mutationActionTypes", "ENSURE_SHEET"),
                            TaskRecipes.ref("workbookPersistenceTypes", "SAVE_AS")),
                        List.of("Author the destination shape before adding later visual layers.")),
                    TaskRecipes.phase(
                        TaskPhasePurpose.VERIFY,
                        "Inspect And Verify",
                        "Read back pivot facts and broader workbook health after authoring.",
                        List.of(
                            TaskRecipes.ref("inspectionQueryTypes", "GET_PIVOT_TABLES"),
                            TaskRecipes.ref("inspectionQueryTypes", "ANALYZE_WORKBOOK_FINDINGS"),
                            TaskRecipes.ref("assertionTypes", "EXPECT_PIVOT_TABLE_PRESENT")),
                        List.of("Use pivot readback to confirm the authored summary surface."))),
                List.of(
                    "Pivot authoring depends on stable source headers and range shapes.",
                    "Calculated fields and unsupported advanced pivot features may require later manual workbook refinement.",
                    "If the workflow grows visual layers, keep the pivot and chart responsibilities separate."))),
        starterPlan());
  }

  private static WorkbookPlan starterPlan() {
    String taskId = "PIVOT_REPORT";
    return ExampleWorkbookPlans.defaultExecutionPlan(
        TaskStarters.taskPlanId(taskId),
        new WorkbookPlan.WorkbookSource.New(),
        ExampleWorkbookPlans.saveAs(TaskStarters.taskWorkbookPath(taskId)),
        ExampleSteps.step(
            new StepId("ensure-data"),
            ExampleSelectors.sheet("Data"),
            new WorkbookMutationAction.EnsureSheet()),
        ExampleSteps.step(
            new StepId("ensure-report"),
            ExampleSelectors.sheet("RangeReport"),
            new WorkbookMutationAction.EnsureSheet()),
        ExampleSteps.step(
            new StepId("seed-pivot-range"),
            ExampleSelectors.range("Data", "A1:D5"),
            new CellMutationAction.SetRange(
                ExampleCellValues.rows(
                    ExampleCellValues.row(
                        ExampleCellValues.text("Region"),
                        ExampleCellValues.text("Stage"),
                        ExampleCellValues.text("Owner"),
                        ExampleCellValues.text("Amount")),
                    ExampleCellValues.row(
                        ExampleCellValues.text("North"),
                        ExampleCellValues.text("Plan"),
                        ExampleCellValues.text("Ada"),
                        ExampleCellValues.number(10.0d)),
                    ExampleCellValues.row(
                        ExampleCellValues.text("North"),
                        ExampleCellValues.text("Do"),
                        ExampleCellValues.text("Ada"),
                        ExampleCellValues.number(15.0d)),
                    ExampleCellValues.row(
                        ExampleCellValues.text("South"),
                        ExampleCellValues.text("Plan"),
                        ExampleCellValues.text("Lin"),
                        ExampleCellValues.number(7.0d)),
                    ExampleCellValues.row(
                        ExampleCellValues.text("South"),
                        ExampleCellValues.text("Do"),
                        ExampleCellValues.text("Lin"),
                        ExampleCellValues.number(12.0d))))),
        ExampleSteps.step(
            new StepId("set-source-table"),
            ExampleSelectors.table("PivotSource", "Data"),
            new StructuredMutationAction.SetTable(
                TableSpec.withDefaultMetadata(
                    new TableName("PivotSource"),
                    new SheetName("Data"),
                    new RangeRef("A1:D5"),
                    false,
                    new TableStyleSpec.Named("TableStyleMedium9", false, false, true, false)))),
        ExampleSteps.step(
            new StepId("author-pivot"),
            new PivotTableSelector.ByNameOnSheet("RegionalTotals", new SheetName("RangeReport")),
            new StructuredMutationAction.SetPivotTable(
                ExamplePivotInputs.regionalTotalsPivotFromTable(
                    "RegionalTotals", "RangeReport", "PivotSource"))),
        ExampleSteps.read(
            new StepId("read-pivots"),
            new PivotTableSelector.ByName("RegionalTotals"),
            new WorkbookAssetIntrospectionQuery.GetPivotTables()),
        ExampleSteps.assertStep(
            new StepId("assert-pivot-present"),
            new PivotTableSelector.ByName("RegionalTotals"),
            new PresenceAssertion.PivotTablePresent()),
        ExampleSteps.read(
            new StepId("read-pivot-findings"),
            ExampleSelectors.workbook(),
            new InspectionAnalysisQuery.AnalyzeWorkbookFindings()));
  }
}
