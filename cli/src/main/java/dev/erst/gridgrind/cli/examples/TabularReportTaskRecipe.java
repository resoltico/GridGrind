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
import dev.erst.gridgrind.contract.action.StructuredMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.dto.TableInput;
import dev.erst.gridgrind.contract.dto.TableStyleInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.SheetIntrospectionQuery;
import dev.erst.gridgrind.contract.query.WorkbookAssetIntrospectionQuery;
import dev.erst.gridgrind.contract.query.WorkbookIntrospectionQuery;
import java.util.List;

/** Published task recipe for structured worksheet reporting workflows. */
final class TabularReportTaskRecipe {
  private TabularReportTaskRecipe() {}

  static GridGrindTaskRecipeDefinition definition() {
    return GridGrindTaskRecipeSupport.definition(
        task(
            "TABULAR_REPORT",
            selfContainedStarter("TABULAR_REPORT"),
            List.of("office", "reporting", "table", "report", "save"),
            discovery(
                List.of(
                    "report", "spreadsheet", "worksheet", "budget", "tracker", "ledger", "table"),
                intent(
                    List.of(TaskGoalKind.AUTHOR, TaskGoalKind.VERIFY),
                    List.of(
                        TaskArtifactKind.WORKBOOK,
                        TaskArtifactKind.SHEET,
                        TaskArtifactKind.CELL,
                        TaskArtifactKind.TABLE))),
            narrative(
                "Create a structured worksheet report with typed cells, table semantics, and factual readback.",
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
                        List.of(
                            "Use factual readback instead of assuming writes landed correctly."))),
                List.of(
                    "Large authored literals belong in UTF8_FILE, FILE, or STANDARD_INPUT sources instead of huge inline JSON.",
                    "Table headers must remain nonblank and unique.",
                    "Formula authoring is scalar-only unless you use SET_ARRAY_FORMULA."))),
        starterPlan());
  }

  private static WorkbookPlan starterPlan() {
    String taskId = "TABULAR_REPORT";
    return ExampleWorkbookPlans.defaultExecutionPlan(
        TaskStarterRecipeSupport.taskPlanId(taskId),
        new WorkbookPlan.WorkbookSource.New(),
        ExampleWorkbookPlans.saveAs(TaskStarterRecipeSupport.taskWorkbookPath(taskId)),
        ExampleSteps.step(
            "ensure-report",
            ExampleSelectors.sheet("Report"),
            new WorkbookMutationAction.EnsureSheet()),
        ExampleSteps.step(
            "seed-report",
            ExampleSelectors.range("Report", "A1:C4"),
            new CellMutationAction.SetRange(
                ExampleCellValues.rows(
                    ExampleCellValues.row(
                        ExampleCellValues.text("Category"),
                        ExampleCellValues.text("Owner"),
                        ExampleCellValues.text("Amount")),
                    ExampleCellValues.row(
                        ExampleCellValues.text("Travel"),
                        ExampleCellValues.text("Ada"),
                        ExampleCellValues.number(125.0d)),
                    ExampleCellValues.row(
                        ExampleCellValues.text("Software"),
                        ExampleCellValues.text("Lin"),
                        ExampleCellValues.number(310.0d)),
                    ExampleCellValues.row(
                        ExampleCellValues.text("Training"),
                        ExampleCellValues.text("Mia"),
                        ExampleCellValues.number(90.0d))))),
        ExampleSteps.step(
            "set-report-table",
            ExampleSelectors.table("QuarterlyReport", "Report"),
            new StructuredMutationAction.SetTable(
                TableInput.withDefaultMetadata(
                    "QuarterlyReport",
                    "Report",
                    "A1:C4",
                    false,
                    new TableStyleInput.Named("TableStyleMedium2", false, false, true, false)))),
        ExampleSteps.step(
            "auto-size-report",
            ExampleSelectors.sheet("Report"),
            new WorkbookMutationAction.AutoSizeColumns()),
        ExampleSteps.read(
            "read-report-cells",
            ExampleSelectors.cells("Report", "A1", "B2", "C4"),
            new SheetIntrospectionQuery.GetCells()),
        ExampleSteps.read(
            "read-report-tables",
            ExampleSelectors.table("QuarterlyReport", "Report"),
            new WorkbookAssetIntrospectionQuery.GetTables()),
        ExampleSteps.read(
            "read-report-workbook",
            ExampleSelectors.workbook(),
            new WorkbookIntrospectionQuery.GetWorkbookSummary()));
  }
}
