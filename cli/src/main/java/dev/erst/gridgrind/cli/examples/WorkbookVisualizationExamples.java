package dev.erst.gridgrind.cli.examples;

import dev.erst.gridgrind.core.model.drawing.PictureDataInput;
import dev.erst.gridgrind.core.model.drawing.SignatureLineInput;
import dev.erst.gridgrind.core.model.formula.NamedRangeTarget;
import dev.erst.gridgrind.core.plan.ChartSelector;
import dev.erst.gridgrind.core.plan.DrawingMutationAction;
import dev.erst.gridgrind.core.plan.InspectionAnalysisQuery;
import dev.erst.gridgrind.core.plan.NamedRangeSelector;
import dev.erst.gridgrind.core.plan.PivotTableSelector;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookAssetIntrospectionQuery;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.source.BinarySourceInput;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import dev.erst.gridgrind.core.vocabulary.NamedRangeScope;
import dev.erst.gridgrind.core.vocabulary.PlanId;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.util.Optional;

/** Generated examples for visual workbook surfaces such as signatures, charts, and pivots. */
final class WorkbookVisualizationExamples {
  private WorkbookVisualizationExamples() {}

  static WorkbookPlan signatureLinePlan(ExamplePathLayout paths) {
    return ExampleWorkbookPlans.defaultExecutionPlan(
        new PlanId("signature-line-workflow"),
        new WorkbookPlan.WorkbookSource.New(),
        ExampleWorkbookPlans.saveAs(paths.generatedWorkbook("gridgrind-signature-line.xlsx")),
        ExampleSteps.step(
            new StepId("step-01-ensure-sheet"),
            ExampleSelectors.sheet("Approvals"),
            new WorkbookMutationAction.EnsureSheet()),
        ExampleSteps.step(
            new StepId("step-02-set-signature-line"),
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
                            BinarySourceInput.inlineBase64(
                                TaskStarterRecipeSupport.onePixelPngBase64())))))),
        ExampleSteps.read(
            new StepId("step-03-read-drawing-objects"),
            new dev.erst.gridgrind.core.plan.DrawingObjectSelector.AllOnSheet("Approvals"),
            new WorkbookAssetIntrospectionQuery.GetDrawingObjects()),
        ExampleSteps.step(
            new StepId("step-04-move-signature-line"),
            new dev.erst.gridgrind.core.plan.DrawingObjectSelector.ByName(
                "Approvals", "BudgetSignature"),
            new DrawingMutationAction.SetDrawingObjectAnchor(
                ExampleDrawingAnchors.anchor(5, 1, 8, 6))),
        ExampleSteps.read(
            new StepId("step-05-read-drawing-objects-after-move"),
            new dev.erst.gridgrind.core.plan.DrawingObjectSelector.AllOnSheet("Approvals"),
            new WorkbookAssetIntrospectionQuery.GetDrawingObjects()));
  }

  static WorkbookPlan chartPlan(ExamplePathLayout paths) {
    return ExampleWorkbookPlans.defaultExecutionPlan(
        new PlanId("chart-workflow"),
        new WorkbookPlan.WorkbookSource.New(),
        ExampleWorkbookPlans.saveAs(paths.generatedWorkbook("gridgrind-chart.xlsx")),
        ExampleSteps.step(
            new StepId("step-01-ensure-sheet"),
            ExampleSelectors.sheet("Ops"),
            new WorkbookMutationAction.EnsureSheet()),
        ExampleSteps.step(
            new StepId("step-02-set-range"),
            ExampleSelectors.range("Ops", "A1:C4"),
            new dev.erst.gridgrind.core.plan.CellMutationAction.SetRange(
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
            new StepId("step-03-set-categories"),
            new NamedRangeSelector.WorkbookScope("ChartCategories"),
            new StructuredMutationAction.SetNamedRange(
                "ChartCategories",
                new NamedRangeScope.Workbook(),
                NamedRangeTarget.range("Ops", "A2:A4"))),
        ExampleSteps.step(
            new StepId("step-04-set-actual"),
            new NamedRangeSelector.WorkbookScope("ChartActual"),
            new StructuredMutationAction.SetNamedRange(
                "ChartActual",
                new NamedRangeScope.Workbook(),
                NamedRangeTarget.range("Ops", "C2:C4"))),
        ExampleSteps.step(
            new StepId("step-05-set-chart"),
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
            new StepId("charts"),
            new ChartSelector.AllOnSheet("Ops"),
            new WorkbookAssetIntrospectionQuery.GetCharts()),
        ExampleSteps.assertStep(
            new StepId("chart-present"),
            new ChartSelector.ByName("Ops", "OpsChart"),
            new dev.erst.gridgrind.core.plan.PresenceAssertion.ChartPresent()));
  }

  static WorkbookPlan pivotPlan(ExamplePathLayout paths) {
    return ExampleWorkbookPlans.defaultExecutionPlan(
        new PlanId("pivot-workflow"),
        new WorkbookPlan.WorkbookSource.New(),
        ExampleWorkbookPlans.saveAs(paths.generatedWorkbook("gridgrind-pivot.xlsx")),
        ExampleSteps.step(
            new StepId("step-01-ensure-data"),
            ExampleSelectors.sheet("Data"),
            new WorkbookMutationAction.EnsureSheet()),
        ExampleSteps.step(
            new StepId("step-02-ensure-report"),
            ExampleSelectors.sheet("RangeReport"),
            new WorkbookMutationAction.EnsureSheet()),
        ExampleSteps.step(
            new StepId("step-03-set-range"),
            ExampleSelectors.range("Data", "A1:D5"),
            new dev.erst.gridgrind.core.plan.CellMutationAction.SetRange(
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
            new StepId("step-04-set-pivot"),
            new PivotTableSelector.ByNameOnSheet("RegionalTotals", "RangeReport"),
            new StructuredMutationAction.SetPivotTable(
                ExamplePivotInputs.regionalTotalsPivotFromRange(
                    "RegionalTotals", "RangeReport", "Data", "A1:D5"))),
        ExampleSteps.read(
            new StepId("pivot-tables"),
            new PivotTableSelector.ByName("RegionalTotals"),
            new WorkbookAssetIntrospectionQuery.GetPivotTables()),
        ExampleSteps.read(
            new StepId("pivot-health"),
            new PivotTableSelector.ByName("RegionalTotals"),
            new InspectionAnalysisQuery.AnalyzePivotTableHealth()));
  }
}
