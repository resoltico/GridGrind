package dev.erst.gridgrind.cli.examples;

import dev.erst.gridgrind.contract.action.DrawingMutationAction;
import dev.erst.gridgrind.contract.action.StructuredMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.dto.ChartDataSourceInput;
import dev.erst.gridgrind.contract.dto.ChartInput;
import dev.erst.gridgrind.contract.dto.ChartLegendInput;
import dev.erst.gridgrind.contract.dto.ChartPlotInput;
import dev.erst.gridgrind.contract.dto.ChartSeriesInput;
import dev.erst.gridgrind.contract.dto.ChartTitleInput;
import dev.erst.gridgrind.contract.dto.NamedRangeScope;
import dev.erst.gridgrind.contract.dto.NamedRangeTarget;
import dev.erst.gridgrind.contract.dto.PictureDataInput;
import dev.erst.gridgrind.contract.dto.PivotTableInput;
import dev.erst.gridgrind.contract.dto.SignatureLineInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.InspectionAnalysisQuery;
import dev.erst.gridgrind.contract.query.WorkbookAssetIntrospectionQuery;
import dev.erst.gridgrind.contract.selector.ChartSelector;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import dev.erst.gridgrind.contract.source.BinarySourceInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarDirection;
import dev.erst.gridgrind.excel.foundation.ExcelChartBarGrouping;
import dev.erst.gridgrind.excel.foundation.ExcelChartLegendPosition;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import dev.erst.gridgrind.excel.foundation.ExcelPivotDataConsolidateFunction;
import java.util.List;
import java.util.Optional;

/** Generated examples for visual workbook surfaces such as signatures, charts, and pivots. */
final class WorkbookVisualizationExamples {
  private static final String ONE_PIXEL_PNG_BASE64 =
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII=";

  private WorkbookVisualizationExamples() {}

  static GridGrindShippedExamples.ShippedExample signatureLineExample(ExamplePathLayout paths) {
    return ExamplePlanSupport.example(
        "SIGNATURE_LINE",
        "signature-line-request.json",
        "Signature-line authoring with drawing-object readback and authored anchor replacement.",
        ExamplePlanSupport.defaultExecutionPlan(
            "signature-line-workflow",
            new WorkbookPlan.WorkbookSource.New(),
            ExamplePlanSupport.saveAs(paths.generatedWorkbook("gridgrind-signature-line.xlsx")),
            ExamplePlanSupport.step(
                "step-01-ensure-sheet",
                ExamplePlanSupport.sheet("Approvals"),
                new WorkbookMutationAction.EnsureSheet()),
            ExamplePlanSupport.step(
                "step-02-set-signature-line",
                ExamplePlanSupport.sheet("Approvals"),
                new DrawingMutationAction.SetSignatureLine(
                    new SignatureLineInput(
                        "BudgetSignature",
                        ExamplePlanSupport.anchor(1, 1, 4, 6),
                        false,
                        Optional.of("Review the budget before signing."),
                        Optional.of("Ada Lovelace"),
                        Optional.of("Finance"),
                        Optional.of("ada@example.com"),
                        Optional.empty(),
                        Optional.of("invalid"),
                        Optional.of(
                            new PictureDataInput(
                                ExcelPictureFormat.PNG,
                                BinarySourceInput.inlineBase64(ONE_PIXEL_PNG_BASE64)))))),
            ExamplePlanSupport.read(
                "step-03-read-drawing-objects",
                new dev.erst.gridgrind.contract.selector.DrawingObjectSelector.AllOnSheet(
                    "Approvals"),
                new WorkbookAssetIntrospectionQuery.GetDrawingObjects()),
            ExamplePlanSupport.step(
                "step-04-move-signature-line",
                new dev.erst.gridgrind.contract.selector.DrawingObjectSelector.ByName(
                    "Approvals", "BudgetSignature"),
                new DrawingMutationAction.SetDrawingObjectAnchor(
                    ExamplePlanSupport.anchor(5, 1, 8, 6))),
            ExamplePlanSupport.read(
                "step-05-read-drawing-objects-after-move",
                new dev.erst.gridgrind.contract.selector.DrawingObjectSelector.AllOnSheet(
                    "Approvals"),
                new WorkbookAssetIntrospectionQuery.GetDrawingObjects())));
  }

  static GridGrindShippedExamples.ShippedExample chartExample(ExamplePathLayout paths) {
    return ExamplePlanSupport.example(
        "CHART",
        "chart-request.json",
        "Supported chart authoring with named-range-backed series and factual chart readback.",
        ExamplePlanSupport.defaultExecutionPlan(
            "chart-workflow",
            new WorkbookPlan.WorkbookSource.New(),
            ExamplePlanSupport.saveAs(paths.generatedWorkbook("gridgrind-chart.xlsx")),
            ExamplePlanSupport.step(
                "step-01-ensure-sheet",
                ExamplePlanSupport.sheet("Ops"),
                new WorkbookMutationAction.EnsureSheet()),
            ExamplePlanSupport.step(
                "step-02-set-range",
                ExamplePlanSupport.range("Ops", "A1:C4"),
                new dev.erst.gridgrind.contract.action.CellMutationAction.SetRange(
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
                "step-03-set-categories",
                new NamedRangeSelector.WorkbookScope("ChartCategories"),
                new StructuredMutationAction.SetNamedRange(
                    "ChartCategories",
                    new NamedRangeScope.Workbook(),
                    NamedRangeTarget.range("Ops", "A2:A4"))),
            ExamplePlanSupport.step(
                "step-04-set-actual",
                new NamedRangeSelector.WorkbookScope("ChartActual"),
                new StructuredMutationAction.SetNamedRange(
                    "ChartActual",
                    new NamedRangeScope.Workbook(),
                    NamedRangeTarget.range("Ops", "C2:C4"))),
            ExamplePlanSupport.step(
                "step-05-set-chart",
                ExamplePlanSupport.sheet("Ops"),
                new DrawingMutationAction.SetChart(
                    new ChartInput(
                        "OpsChart",
                        ExamplePlanSupport.anchor(4, 0, 8, 12),
                        new ChartTitleInput.Text(TextSourceInput.inline("Roadmap")),
                        new ChartLegendInput.Visible(ExcelChartLegendPosition.TOP_RIGHT),
                        dev.erst.gridgrind.excel.foundation.ExcelChartDisplayBlanksAs.SPAN,
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
                                        new ChartDataSourceInput.Reference("ChartCategories"),
                                        new ChartDataSourceInput.Reference("Ops!$B$2:$B$4"),
                                        Optional.empty(),
                                        Optional.empty(),
                                        Optional.empty(),
                                        Optional.empty()),
                                    new ChartSeriesInput(
                                        new ChartTitleInput.Text(TextSourceInput.inline("Actual")),
                                        new ChartDataSourceInput.Reference("ChartCategories"),
                                        new ChartDataSourceInput.Reference("ChartActual"),
                                        Optional.empty(),
                                        Optional.empty(),
                                        Optional.empty(),
                                        Optional.empty()))))))),
            ExamplePlanSupport.read(
                "charts",
                new ChartSelector.AllOnSheet("Ops"),
                new WorkbookAssetIntrospectionQuery.GetCharts()),
            ExamplePlanSupport.assertStep(
                "chart-present",
                new ChartSelector.ByName("Ops", "OpsChart"),
                new dev.erst.gridgrind.contract.assertion.PresenceAssertion.ChartPresent())));
  }

  static GridGrindShippedExamples.ShippedExample pivotExample(ExamplePathLayout paths) {
    return ExamplePlanSupport.example(
        "PIVOT",
        "pivot-request.json",
        "Pivot authoring from a contiguous range with pivot readback and health analysis.",
        ExamplePlanSupport.defaultExecutionPlan(
            "pivot-workflow",
            new WorkbookPlan.WorkbookSource.New(),
            ExamplePlanSupport.saveAs(paths.generatedWorkbook("gridgrind-pivot.xlsx")),
            ExamplePlanSupport.step(
                "step-01-ensure-data",
                ExamplePlanSupport.sheet("Data"),
                new WorkbookMutationAction.EnsureSheet()),
            ExamplePlanSupport.step(
                "step-02-ensure-report",
                ExamplePlanSupport.sheet("RangeReport"),
                new WorkbookMutationAction.EnsureSheet()),
            ExamplePlanSupport.step(
                "step-03-set-range",
                ExamplePlanSupport.range("Data", "A1:D5"),
                new dev.erst.gridgrind.contract.action.CellMutationAction.SetRange(
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
                "step-04-set-pivot",
                new PivotTableSelector.ByNameOnSheet("RegionalTotals", "RangeReport"),
                new StructuredMutationAction.SetPivotTable(
                    new PivotTableInput(
                        "RegionalTotals",
                        "RangeReport",
                        new PivotTableInput.Source.Range("Data", "A1:D5"),
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
                "pivot-tables",
                new PivotTableSelector.ByName("RegionalTotals"),
                new WorkbookAssetIntrospectionQuery.GetPivotTables()),
            ExamplePlanSupport.read(
                "pivot-health",
                new PivotTableSelector.ByName("RegionalTotals"),
                new InspectionAnalysisQuery.AnalyzePivotTableHealth())));
  }
}
