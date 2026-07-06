package dev.erst.gridgrind.cli.examples;

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
import dev.erst.gridgrind.contract.action.StructuredMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.assertion.PresenceAssertion;
import dev.erst.gridgrind.contract.dto.TableInput;
import dev.erst.gridgrind.contract.dto.TableStyleInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.InspectionAnalysisQuery;
import dev.erst.gridgrind.contract.query.WorkbookAssetIntrospectionQuery;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import java.util.List;

/** Published task recipe for pivot-table reporting workflows. */
final class PivotReportTaskRecipe {
  private PivotReportTaskRecipe() {}

  static GridGrindTaskRecipeDefinition definition() {
    return GridGrindTaskRecipeSupport.definition(
        GridGrindTaskRecipeSupport.task(
            "PIVOT_REPORT",
            GridGrindTaskRecipeSupport.selfContainedStarter("PIVOT_REPORT"),
            List.of("office", "pivot", "summary", "analysis", "table"),
            GridGrindTaskRecipeSupport.discovery(
                List.of("pivot", "pivot table", "summary", "analysis", "cross-tab"),
                GridGrindTaskRecipeSupport.intent(
                    List.of(TaskGoalKind.AUTHOR, TaskGoalKind.VERIFY, TaskGoalKind.ANALYZE),
                    List.of(
                        TaskArtifactKind.WORKBOOK,
                        TaskArtifactKind.TABLE,
                        TaskArtifactKind.PIVOT_TABLE))),
            GridGrindTaskRecipeSupport.narrative(
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
            GridGrindTaskRecipeSupport.profile(
                TaskSourceMode.NEW_WORKBOOK,
                TaskPersistenceMode.SAVE_AS,
                TaskMutationMode.MUTATING,
                TaskAssetMode.SELF_CONTAINED),
            GridGrindTaskRecipeSupport.signals(
                List.of(
                    TaskInputKind.TARGET_SHEET_NAMES,
                    TaskInputKind.TARGET_OBJECT_NAMES,
                    TaskInputKind.TABULAR_SOURCE_ROWS,
                    TaskInputKind.PERSISTENCE_TARGET_PATH),
                List.of(
                    TaskVerificationKind.FACT_READBACK,
                    TaskVerificationKind.ASSERTION_CHECKS,
                    TaskVerificationKind.HEALTH_ANALYSIS)),
            GridGrindTaskRecipeSupport.workflow(
                List.of(
                    GridGrindTaskRecipeSupport.phase(
                        TaskPhasePurpose.PREPARE,
                        "Model Source Data",
                        "Create one table-backed source surface for later summarization.",
                        List.of(
                            GridGrindTaskRecipeSupport.ref("sourceTypes", "NEW"),
                            GridGrindTaskRecipeSupport.ref("mutationActionTypes", "ENSURE_SHEET"),
                            GridGrindTaskRecipeSupport.ref("mutationActionTypes", "SET_RANGE"),
                            GridGrindTaskRecipeSupport.ref("mutationActionTypes", "SET_TABLE")),
                        List.of("Pivot work starts from explicit tabular semantics.")),
                    GridGrindTaskRecipeSupport.phase(
                        TaskPhasePurpose.AUTHOR,
                        "Author The Pivot",
                        "Create the pivot table and place it on a durable destination sheet.",
                        List.of(
                            GridGrindTaskRecipeSupport.ref(
                                "mutationActionTypes", "SET_PIVOT_TABLE"),
                            GridGrindTaskRecipeSupport.ref("mutationActionTypes", "ENSURE_SHEET"),
                            GridGrindTaskRecipeSupport.ref("persistenceTypes", "SAVE_AS")),
                        List.of("Author the destination shape before adding later visual layers.")),
                    GridGrindTaskRecipeSupport.phase(
                        TaskPhasePurpose.VERIFY,
                        "Inspect And Verify",
                        "Read back pivot facts and broader workbook health after authoring.",
                        List.of(
                            GridGrindTaskRecipeSupport.ref(
                                "inspectionQueryTypes", "GET_PIVOT_TABLES"),
                            GridGrindTaskRecipeSupport.ref(
                                "inspectionQueryTypes", "ANALYZE_WORKBOOK_FINDINGS"),
                            GridGrindTaskRecipeSupport.ref(
                                "assertionTypes", "EXPECT_PIVOT_TABLE_PRESENT")),
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
        TaskStarterRecipeSupport.taskPlanId(taskId),
        new WorkbookPlan.WorkbookSource.New(),
        ExampleWorkbookPlans.saveAs(TaskStarterRecipeSupport.taskWorkbookPath(taskId)),
        ExampleSteps.step(
            "ensure-data",
            ExampleSelectors.sheet("Data"),
            new WorkbookMutationAction.EnsureSheet()),
        ExampleSteps.step(
            "ensure-report",
            ExampleSelectors.sheet("RangeReport"),
            new WorkbookMutationAction.EnsureSheet()),
        ExampleSteps.step(
            "seed-pivot-range",
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
            "set-source-table",
            ExampleSelectors.table("PivotSource", "Data"),
            new StructuredMutationAction.SetTable(
                TableInput.withDefaultMetadata(
                    "PivotSource",
                    "Data",
                    "A1:D5",
                    false,
                    new TableStyleInput.Named("TableStyleMedium9", false, false, true, false)))),
        ExampleSteps.step(
            "author-pivot",
            new PivotTableSelector.ByNameOnSheet("RegionalTotals", "RangeReport"),
            new StructuredMutationAction.SetPivotTable(
                ExamplePivotInputs.regionalTotalsPivotFromTable(
                    "RegionalTotals", "RangeReport", "PivotSource"))),
        ExampleSteps.read(
            "read-pivots",
            new PivotTableSelector.ByName("RegionalTotals"),
            new WorkbookAssetIntrospectionQuery.GetPivotTables()),
        ExampleSteps.assertStep(
            "assert-pivot-present",
            new PivotTableSelector.ByName("RegionalTotals"),
            new PresenceAssertion.PivotTablePresent()),
        ExampleSteps.read(
            "read-pivot-findings",
            ExampleSelectors.workbook(),
            new InspectionAnalysisQuery.AnalyzeWorkbookFindings()));
  }
}
