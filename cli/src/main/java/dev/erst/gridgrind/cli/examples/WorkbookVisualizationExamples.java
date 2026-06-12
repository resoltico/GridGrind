package dev.erst.gridgrind.cli.examples;

import dev.erst.gridgrind.contract.action.DrawingMutationAction;
import dev.erst.gridgrind.contract.action.StructuredMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.dto.NamedRangeScope;
import dev.erst.gridgrind.contract.dto.NamedRangeTarget;
import dev.erst.gridgrind.contract.dto.PictureDataInput;
import dev.erst.gridgrind.contract.dto.SignatureLineInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.InspectionAnalysisQuery;
import dev.erst.gridgrind.contract.query.WorkbookAssetIntrospectionQuery;
import dev.erst.gridgrind.contract.selector.ChartSelector;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import dev.erst.gridgrind.contract.source.BinarySourceInput;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import java.util.Optional;

/** Generated examples for visual workbook surfaces such as signatures, charts, and pivots. */
final class WorkbookVisualizationExamples {
  private static final String ONE_PIXEL_PNG_BASE64 =
      "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII=";

  private WorkbookVisualizationExamples() {}

  static GridGrindShippedExamples.ShippedExample signatureLineExample(ExamplePathLayout paths) {
    return ExampleDefinitions.example(
        "SIGNATURE_LINE",
        "signature-line-request.json",
        "Signature-line authoring with drawing-object readback and authored anchor replacement.",
        ExampleWorkbookPlans.defaultExecutionPlan(
            "signature-line-workflow",
            new WorkbookPlan.WorkbookSource.New(),
            ExampleWorkbookPlans.saveAs(paths.generatedWorkbook("gridgrind-signature-line.xlsx")),
            ExampleSteps.step(
                "step-01-ensure-sheet",
                ExampleSelectors.sheet("Approvals"),
                new WorkbookMutationAction.EnsureSheet()),
            ExampleSteps.step(
                "step-02-set-signature-line",
                ExampleSelectors.sheet("Approvals"),
                new DrawingMutationAction.SetSignatureLine(
                    new SignatureLineInput(
                        "BudgetSignature",
                        ExampleDrawingAnchors.anchor(1, 1, 4, 6),
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
            ExampleSteps.read(
                "step-03-read-drawing-objects",
                new dev.erst.gridgrind.contract.selector.DrawingObjectSelector.AllOnSheet(
                    "Approvals"),
                new WorkbookAssetIntrospectionQuery.GetDrawingObjects()),
            ExampleSteps.step(
                "step-04-move-signature-line",
                new dev.erst.gridgrind.contract.selector.DrawingObjectSelector.ByName(
                    "Approvals", "BudgetSignature"),
                new DrawingMutationAction.SetDrawingObjectAnchor(
                    ExampleDrawingAnchors.anchor(5, 1, 8, 6))),
            ExampleSteps.read(
                "step-05-read-drawing-objects-after-move",
                new dev.erst.gridgrind.contract.selector.DrawingObjectSelector.AllOnSheet(
                    "Approvals"),
                new WorkbookAssetIntrospectionQuery.GetDrawingObjects())));
  }

  static GridGrindShippedExamples.ShippedExample chartExample(ExamplePathLayout paths) {
    return ExampleDefinitions.example(
        "CHART",
        "chart-request.json",
        "Supported chart authoring with named-range-backed series and factual chart readback.",
        ExampleWorkbookPlans.defaultExecutionPlan(
            "chart-workflow",
            new WorkbookPlan.WorkbookSource.New(),
            ExampleWorkbookPlans.saveAs(paths.generatedWorkbook("gridgrind-chart.xlsx")),
            ExampleSteps.step(
                "step-01-ensure-sheet",
                ExampleSelectors.sheet("Ops"),
                new WorkbookMutationAction.EnsureSheet()),
            ExampleSteps.step(
                "step-02-set-range",
                ExampleSelectors.range("Ops", "A1:C4"),
                new dev.erst.gridgrind.contract.action.CellMutationAction.SetRange(
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
                "step-03-set-categories",
                new NamedRangeSelector.WorkbookScope("ChartCategories"),
                new StructuredMutationAction.SetNamedRange(
                    "ChartCategories",
                    new NamedRangeScope.Workbook(),
                    NamedRangeTarget.range("Ops", "A2:A4"))),
            ExampleSteps.step(
                "step-04-set-actual",
                new NamedRangeSelector.WorkbookScope("ChartActual"),
                new StructuredMutationAction.SetNamedRange(
                    "ChartActual",
                    new NamedRangeScope.Workbook(),
                    NamedRangeTarget.range("Ops", "C2:C4"))),
            ExampleSteps.step(
                "step-05-set-chart",
                ExampleSelectors.sheet("Ops"),
                new DrawingMutationAction.SetChart(
                    ExampleChartInputs.clusteredColumnComparisonChart(
                        "OpsChart",
                        ExampleDrawingAnchors.anchor(4, 0, 8, 12),
                        "Roadmap",
                        "ChartCategories",
                        "Plan",
                        "Ops!$B$2:$B$4",
                        "Actual",
                        "ChartActual"))),
            ExampleSteps.read(
                "charts",
                new ChartSelector.AllOnSheet("Ops"),
                new WorkbookAssetIntrospectionQuery.GetCharts()),
            ExampleSteps.assertStep(
                "chart-present",
                new ChartSelector.ByName("Ops", "OpsChart"),
                new dev.erst.gridgrind.contract.assertion.PresenceAssertion.ChartPresent())));
  }

  static GridGrindShippedExamples.ShippedExample pivotExample(ExamplePathLayout paths) {
    return ExampleDefinitions.example(
        "PIVOT",
        "pivot-request.json",
        "Pivot authoring from a contiguous range with pivot readback and health analysis.",
        ExampleWorkbookPlans.defaultExecutionPlan(
            "pivot-workflow",
            new WorkbookPlan.WorkbookSource.New(),
            ExampleWorkbookPlans.saveAs(paths.generatedWorkbook("gridgrind-pivot.xlsx")),
            ExampleSteps.step(
                "step-01-ensure-data",
                ExampleSelectors.sheet("Data"),
                new WorkbookMutationAction.EnsureSheet()),
            ExampleSteps.step(
                "step-02-ensure-report",
                ExampleSelectors.sheet("RangeReport"),
                new WorkbookMutationAction.EnsureSheet()),
            ExampleSteps.step(
                "step-03-set-range",
                ExampleSelectors.range("Data", "A1:D5"),
                new dev.erst.gridgrind.contract.action.CellMutationAction.SetRange(
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
                "step-04-set-pivot",
                new PivotTableSelector.ByNameOnSheet("RegionalTotals", "RangeReport"),
                new StructuredMutationAction.SetPivotTable(
                    ExamplePivotInputs.regionalTotalsPivotFromRange(
                        "RegionalTotals", "RangeReport", "Data", "A1:D5"))),
            ExampleSteps.read(
                "pivot-tables",
                new PivotTableSelector.ByName("RegionalTotals"),
                new WorkbookAssetIntrospectionQuery.GetPivotTables()),
            ExampleSteps.read(
                "pivot-health",
                new PivotTableSelector.ByName("RegionalTotals"),
                new InspectionAnalysisQuery.AnalyzePivotTableHealth())));
  }
}
