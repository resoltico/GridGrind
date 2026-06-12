package dev.erst.gridgrind.cli.examples;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.DrawingMutationAction;
import dev.erst.gridgrind.contract.action.StructuredMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.assertion.PresenceAssertion;
import dev.erst.gridgrind.contract.dto.TableInput;
import dev.erst.gridgrind.contract.dto.TableStyleInput;
import dev.erst.gridgrind.contract.query.InspectionAnalysisQuery;
import dev.erst.gridgrind.contract.query.SheetIntrospectionQuery;
import dev.erst.gridgrind.contract.query.WorkbookAssetIntrospectionQuery;
import dev.erst.gridgrind.contract.query.WorkbookIntrospectionQuery;
import dev.erst.gridgrind.contract.selector.ChartSelector;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import java.util.List;

/** Published starter plans for report- and dashboard-style workflows. */
final class GridGrindTaskStarterReportPlans {
  private GridGrindTaskStarterReportPlans() {}

  static List<TaskStarterPlan> starters() {
    return List.of(tabularReportStarter(), dashboardStarter(), pivotReportStarter());
  }

  private static TaskStarterPlan tabularReportStarter() {
    String taskId = "TABULAR_REPORT";
    return TaskStarterPlanSupport.selfContainedStarter(
        taskId,
        ExampleWorkbookPlans.defaultExecutionPlan(
            TaskStarterPlanSupport.taskPlanId(taskId),
            new dev.erst.gridgrind.contract.dto.WorkbookPlan.WorkbookSource.New(),
            ExampleWorkbookPlans.saveAs(TaskStarterPlanSupport.taskWorkbookPath(taskId)),
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
                        new TableStyleInput.Named(
                            "TableStyleMedium2", false, false, true, false)))),
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
                new WorkbookIntrospectionQuery.GetWorkbookSummary())));
  }

  private static TaskStarterPlan dashboardStarter() {
    String taskId = "DASHBOARD";
    return TaskStarterPlanSupport.selfContainedStarter(
        taskId,
        ExampleWorkbookPlans.defaultExecutionPlan(
            TaskStarterPlanSupport.taskPlanId(taskId),
            new dev.erst.gridgrind.contract.dto.WorkbookPlan.WorkbookSource.New(),
            ExampleWorkbookPlans.saveAs(TaskStarterPlanSupport.taskWorkbookPath(taskId)),
            ExampleSteps.step(
                "ensure-ops",
                ExampleSelectors.sheet("Ops"),
                new WorkbookMutationAction.EnsureSheet()),
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
                new ChartSelector.AllOnSheet("Ops"),
                new WorkbookAssetIntrospectionQuery.GetCharts()),
            ExampleSteps.assertStep(
                "assert-dashboard-chart",
                new ChartSelector.ByName("Ops", "DashboardChart"),
                new PresenceAssertion.ChartPresent())));
  }

  private static TaskStarterPlan pivotReportStarter() {
    String taskId = "PIVOT_REPORT";
    return TaskStarterPlanSupport.selfContainedStarter(
        taskId,
        ExampleWorkbookPlans.defaultExecutionPlan(
            TaskStarterPlanSupport.taskPlanId(taskId),
            new dev.erst.gridgrind.contract.dto.WorkbookPlan.WorkbookSource.New(),
            ExampleWorkbookPlans.saveAs(TaskStarterPlanSupport.taskWorkbookPath(taskId)),
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
                        new TableStyleInput.Named(
                            "TableStyleMedium9", false, false, true, false)))),
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
                new InspectionAnalysisQuery.AnalyzeWorkbookFindings())));
  }
}
