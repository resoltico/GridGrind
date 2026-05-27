package dev.erst.gridgrind.cli.examples;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.DrawingMutationAction;
import dev.erst.gridgrind.contract.action.StructuredMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.assertion.PresenceAssertion;
import dev.erst.gridgrind.contract.dto.ChartDataSourceInput;
import dev.erst.gridgrind.contract.dto.ChartInput;
import dev.erst.gridgrind.contract.dto.ChartLegendInput;
import dev.erst.gridgrind.contract.dto.ChartPlotInput;
import dev.erst.gridgrind.contract.dto.ChartSeriesInput;
import dev.erst.gridgrind.contract.dto.ChartTitleInput;
import dev.erst.gridgrind.contract.dto.PivotTableInput;
import dev.erst.gridgrind.contract.dto.TableInput;
import dev.erst.gridgrind.contract.dto.TableStyleInput;
import dev.erst.gridgrind.contract.query.InspectionAnalysisQuery;
import dev.erst.gridgrind.contract.query.SheetIntrospectionQuery;
import dev.erst.gridgrind.contract.query.WorkbookAssetIntrospectionQuery;
import dev.erst.gridgrind.contract.query.WorkbookIntrospectionQuery;
import dev.erst.gridgrind.contract.selector.ChartSelector;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarDirection;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarGrouping;
import dev.erst.gridgrind.excel.foundation.ExcelChartDisplayBlanksAs;
import dev.erst.gridgrind.excel.foundation.ExcelChartLegendPosition;
import dev.erst.gridgrind.excel.foundation.ExcelPivotDataConsolidateFunction;
import java.util.List;
import java.util.Optional;

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
        ExamplePlanSupport.defaultExecutionPlan(
            TaskStarterPlanSupport.taskPlanId(taskId),
            new dev.erst.gridgrind.contract.dto.WorkbookPlan.WorkbookSource.New(),
            ExamplePlanSupport.saveAs(TaskStarterPlanSupport.taskWorkbookPath(taskId)),
            ExamplePlanSupport.step(
                "ensure-report",
                ExamplePlanSupport.sheet("Report"),
                new WorkbookMutationAction.EnsureSheet()),
            ExamplePlanSupport.step(
                "seed-report",
                ExamplePlanSupport.range("Report", "A1:C4"),
                new CellMutationAction.SetRange(
                    ExamplePlanSupport.rows(
                        ExamplePlanSupport.row(
                            ExamplePlanSupport.text("Category"),
                            ExamplePlanSupport.text("Owner"),
                            ExamplePlanSupport.text("Amount")),
                        ExamplePlanSupport.row(
                            ExamplePlanSupport.text("Travel"),
                            ExamplePlanSupport.text("Ada"),
                            ExamplePlanSupport.number(125.0d)),
                        ExamplePlanSupport.row(
                            ExamplePlanSupport.text("Software"),
                            ExamplePlanSupport.text("Lin"),
                            ExamplePlanSupport.number(310.0d)),
                        ExamplePlanSupport.row(
                            ExamplePlanSupport.text("Training"),
                            ExamplePlanSupport.text("Mia"),
                            ExamplePlanSupport.number(90.0d))))),
            ExamplePlanSupport.step(
                "set-report-table",
                ExamplePlanSupport.table("QuarterlyReport", "Report"),
                new StructuredMutationAction.SetTable(
                    TableInput.withDefaultMetadata(
                        "QuarterlyReport",
                        "Report",
                        "A1:C4",
                        false,
                        new TableStyleInput.Named(
                            "TableStyleMedium2", false, false, true, false)))),
            ExamplePlanSupport.step(
                "auto-size-report",
                ExamplePlanSupport.sheet("Report"),
                new WorkbookMutationAction.AutoSizeColumns()),
            ExamplePlanSupport.read(
                "read-report-cells",
                ExamplePlanSupport.cells("Report", "A1", "B2", "C4"),
                new SheetIntrospectionQuery.GetCells()),
            ExamplePlanSupport.read(
                "read-report-tables",
                ExamplePlanSupport.table("QuarterlyReport", "Report"),
                new WorkbookAssetIntrospectionQuery.GetTables()),
            ExamplePlanSupport.read(
                "read-report-workbook",
                ExamplePlanSupport.workbook(),
                new WorkbookIntrospectionQuery.GetWorkbookSummary())));
  }

  private static TaskStarterPlan dashboardStarter() {
    String taskId = "DASHBOARD";
    return TaskStarterPlanSupport.selfContainedStarter(
        taskId,
        ExamplePlanSupport.defaultExecutionPlan(
            TaskStarterPlanSupport.taskPlanId(taskId),
            new dev.erst.gridgrind.contract.dto.WorkbookPlan.WorkbookSource.New(),
            ExamplePlanSupport.saveAs(TaskStarterPlanSupport.taskWorkbookPath(taskId)),
            ExamplePlanSupport.step(
                "ensure-ops",
                ExamplePlanSupport.sheet("Ops"),
                new WorkbookMutationAction.EnsureSheet()),
            ExamplePlanSupport.step(
                "seed-ops-range",
                ExamplePlanSupport.range("Ops", "A1:C4"),
                new CellMutationAction.SetRange(
                    ExamplePlanSupport.rows(
                        ExamplePlanSupport.row(
                            ExamplePlanSupport.text("Month"),
                            ExamplePlanSupport.text("Plan"),
                            ExamplePlanSupport.text("Actual")),
                        ExamplePlanSupport.row(
                            ExamplePlanSupport.text("Jan"),
                            ExamplePlanSupport.number(10.0d),
                            ExamplePlanSupport.number(12.0d)),
                        ExamplePlanSupport.row(
                            ExamplePlanSupport.text("Feb"),
                            ExamplePlanSupport.number(18.0d),
                            ExamplePlanSupport.number(17.0d)),
                        ExamplePlanSupport.row(
                            ExamplePlanSupport.text("Mar"),
                            ExamplePlanSupport.number(15.0d),
                            ExamplePlanSupport.number(16.0d))))),
            ExamplePlanSupport.step(
                "define-chart-categories",
                new NamedRangeSelector.WorkbookScope("DashboardCategories"),
                new StructuredMutationAction.SetNamedRange(
                    "DashboardCategories",
                    new dev.erst.gridgrind.contract.dto.NamedRangeScope.Workbook(),
                    dev.erst.gridgrind.contract.dto.NamedRangeTarget.range("Ops", "A2:A4"))),
            ExamplePlanSupport.step(
                "define-dashboard-actual",
                new NamedRangeSelector.WorkbookScope("DashboardActual"),
                new StructuredMutationAction.SetNamedRange(
                    "DashboardActual",
                    new dev.erst.gridgrind.contract.dto.NamedRangeScope.Workbook(),
                    dev.erst.gridgrind.contract.dto.NamedRangeTarget.range("Ops", "C2:C4"))),
            ExamplePlanSupport.step(
                "author-dashboard-chart",
                ExamplePlanSupport.sheet("Ops"),
                new DrawingMutationAction.SetChart(
                    new ChartInput(
                        "DashboardChart",
                        ExamplePlanSupport.anchor(4, 0, 8, 12),
                        new ChartTitleInput.Text(TextSourceInput.inline("Operations Dashboard")),
                        new ChartLegendInput.Visible(ExcelChartLegendPosition.TOP_RIGHT),
                        ExcelChartDisplayBlanksAs.SPAN,
                        false,
                        List.of(
                            new ChartPlotInput.Bar(
                                true,
                                ExcelChartBarDirection.COLUMN,
                                ExcelChartBarGrouping.CLUSTERED,
                                Optional.of(150),
                                Optional.of(0),
                                List.of(
                                    new ChartSeriesInput(
                                        new ChartTitleInput.Text(TextSourceInput.inline("Plan")),
                                        new ChartDataSourceInput.Reference("DashboardCategories"),
                                        new ChartDataSourceInput.Reference("Ops!$B$2:$B$4"),
                                        Optional.empty(),
                                        Optional.empty(),
                                        Optional.empty(),
                                        Optional.empty()),
                                    new ChartSeriesInput(
                                        new ChartTitleInput.Text(TextSourceInput.inline("Actual")),
                                        new ChartDataSourceInput.Reference("DashboardCategories"),
                                        new ChartDataSourceInput.Reference("DashboardActual"),
                                        Optional.empty(),
                                        Optional.empty(),
                                        Optional.empty(),
                                        Optional.empty()))))))),
            ExamplePlanSupport.read(
                "read-dashboard-charts",
                new ChartSelector.AllOnSheet("Ops"),
                new WorkbookAssetIntrospectionQuery.GetCharts()),
            ExamplePlanSupport.assertStep(
                "assert-dashboard-chart",
                new ChartSelector.ByName("Ops", "DashboardChart"),
                new PresenceAssertion.ChartPresent())));
  }

  private static TaskStarterPlan pivotReportStarter() {
    String taskId = "PIVOT_REPORT";
    return TaskStarterPlanSupport.selfContainedStarter(
        taskId,
        ExamplePlanSupport.defaultExecutionPlan(
            TaskStarterPlanSupport.taskPlanId(taskId),
            new dev.erst.gridgrind.contract.dto.WorkbookPlan.WorkbookSource.New(),
            ExamplePlanSupport.saveAs(TaskStarterPlanSupport.taskWorkbookPath(taskId)),
            ExamplePlanSupport.step(
                "ensure-data",
                ExamplePlanSupport.sheet("Data"),
                new WorkbookMutationAction.EnsureSheet()),
            ExamplePlanSupport.step(
                "ensure-report",
                ExamplePlanSupport.sheet("RangeReport"),
                new WorkbookMutationAction.EnsureSheet()),
            ExamplePlanSupport.step(
                "seed-pivot-range",
                ExamplePlanSupport.range("Data", "A1:D5"),
                new CellMutationAction.SetRange(
                    ExamplePlanSupport.rows(
                        ExamplePlanSupport.row(
                            ExamplePlanSupport.text("Region"),
                            ExamplePlanSupport.text("Stage"),
                            ExamplePlanSupport.text("Owner"),
                            ExamplePlanSupport.text("Amount")),
                        ExamplePlanSupport.row(
                            ExamplePlanSupport.text("North"),
                            ExamplePlanSupport.text("Plan"),
                            ExamplePlanSupport.text("Ada"),
                            ExamplePlanSupport.number(10.0d)),
                        ExamplePlanSupport.row(
                            ExamplePlanSupport.text("North"),
                            ExamplePlanSupport.text("Do"),
                            ExamplePlanSupport.text("Ada"),
                            ExamplePlanSupport.number(15.0d)),
                        ExamplePlanSupport.row(
                            ExamplePlanSupport.text("South"),
                            ExamplePlanSupport.text("Plan"),
                            ExamplePlanSupport.text("Lin"),
                            ExamplePlanSupport.number(7.0d)),
                        ExamplePlanSupport.row(
                            ExamplePlanSupport.text("South"),
                            ExamplePlanSupport.text("Do"),
                            ExamplePlanSupport.text("Lin"),
                            ExamplePlanSupport.number(12.0d))))),
            ExamplePlanSupport.step(
                "set-source-table",
                ExamplePlanSupport.table("PivotSource", "Data"),
                new StructuredMutationAction.SetTable(
                    TableInput.withDefaultMetadata(
                        "PivotSource",
                        "Data",
                        "A1:D5",
                        false,
                        new TableStyleInput.Named(
                            "TableStyleMedium9", false, false, true, false)))),
            ExamplePlanSupport.step(
                "author-pivot",
                new PivotTableSelector.ByNameOnSheet("RegionalTotals", "RangeReport"),
                new StructuredMutationAction.SetPivotTable(
                    new PivotTableInput(
                        "RegionalTotals",
                        "RangeReport",
                        new PivotTableInput.Source.Table("PivotSource"),
                        new PivotTableInput.Anchor("A3"),
                        List.of("Region"),
                        List.of("Stage"),
                        List.of(),
                        List.of(
                            new PivotTableInput.DataField(
                                "Amount",
                                ExcelPivotDataConsolidateFunction.SUM,
                                "Total Amount",
                                Optional.of("#,##0.00")))))),
            ExamplePlanSupport.read(
                "read-pivots",
                new PivotTableSelector.ByName("RegionalTotals"),
                new WorkbookAssetIntrospectionQuery.GetPivotTables()),
            ExamplePlanSupport.assertStep(
                "assert-pivot-present",
                new PivotTableSelector.ByName("RegionalTotals"),
                new PresenceAssertion.PivotTablePresent()),
            ExamplePlanSupport.read(
                "read-pivot-findings",
                ExamplePlanSupport.workbook(),
                new InspectionAnalysisQuery.AnalyzeWorkbookFindings())));
  }
}
