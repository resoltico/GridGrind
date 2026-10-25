package dev.erst.gridgrind.cli.examples;

import dev.erst.gridgrind.core.model.drawing.CommentInput;
import dev.erst.gridgrind.core.model.formula.CalculationPolicyInput;
import dev.erst.gridgrind.core.model.formula.CalculationStrategyInput;
import dev.erst.gridgrind.core.model.formula.NamedRangeTarget;
import dev.erst.gridgrind.core.model.value.HyperlinkTarget;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.ExecutionModeInput;
import dev.erst.gridgrind.core.plan.ExecutionPolicyInput;
import dev.erst.gridgrind.core.plan.InspectionAnalysisQuery;
import dev.erst.gridgrind.core.plan.InspectionSurfaceQuery;
import dev.erst.gridgrind.core.plan.NamedRangeSelector;
import dev.erst.gridgrind.core.plan.SheetIntrospectionQuery;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.source.TextSourceInput;
import dev.erst.gridgrind.core.vocabulary.NamedRangeScope;
import dev.erst.gridgrind.core.vocabulary.PlanId;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;

/** Generated examples for inspection-heavy, health-focused, and execution-mode workflows. */
final class WorkbookAuditExamples {
  private WorkbookAuditExamples() {}

  static WorkbookPlan workbookHealthPlan(ExamplePathLayout paths) {
    return ExampleWorkbookPlans.defaultExecutionPlan(
        new PlanId("workbook-health-workflow"),
        new WorkbookPlan.WorkbookSource.New(),
        new WorkbookPlan.WorkbookPersistence.None(),
        ExampleSteps.step(
            new StepId("step-01-ensure-budget-review"),
            ExampleSelectors.sheet("Budget Review"),
            new WorkbookMutationAction.EnsureSheet()),
        ExampleSteps.step(
            new StepId("step-02-ensure-summary"),
            ExampleSelectors.sheet("Summary"),
            new WorkbookMutationAction.EnsureSheet()),
        ExampleSteps.step(
            new StepId("step-03-set-budget-header"),
            ExampleSelectors.cell("Budget Review", "A1"),
            new CellMutationAction.SetCell(ExampleCellValues.text("Amount"))),
        ExampleSteps.step(
            new StepId("step-04-set-budget-value"),
            ExampleSelectors.cell("Budget Review", "B1"),
            new CellMutationAction.SetCell(ExampleCellValues.number(1200.0d))),
        ExampleSteps.step(
            new StepId("step-05-set-summary-formula"),
            ExampleSelectors.cell("Summary", "A1"),
            new CellMutationAction.SetCell(ExampleCellValues.formula("'Budget Review'!B1"))),
        ExampleSteps.read(
            new StepId("summary-sheet"),
            ExampleSelectors.sheet("Summary"),
            new SheetIntrospectionQuery.GetSheetSummary()),
        ExampleSteps.read(
            new StepId("formula-health"),
            ExampleSelectors.sheets("Summary"),
            new InspectionAnalysisQuery.AnalyzeFormulaHealth()),
        ExampleSteps.read(
            new StepId("lint"),
            ExampleSelectors.workbook(),
            new InspectionAnalysisQuery.AnalyzeWorkbookFindings()),
        ExampleSteps.read(
            new StepId("summary-cells"),
            ExampleSelectors.cells("Summary", "A1"),
            new SheetIntrospectionQuery.GetCells()));
  }

  static WorkbookPlan largeFileModesPlan(ExamplePathLayout paths) {
    return ExampleWorkbookPlans.plan(
        new PlanId("large-file-modes-workflow"),
        new WorkbookPlan.WorkbookSource.New(),
        ExampleWorkbookPlans.saveAs(paths.generatedWorkbook("gridgrind-large-file-modes.xlsx")),
        new ExecutionPolicyInput(
            ExecutionModeInput.streamingWrite(),
            dev.erst.gridgrind.core.plan.ExecutionJournalInput.defaults(),
            new CalculationPolicyInput(new CalculationStrategyInput.DoNotCalculate(), true),
            dev.erst.gridgrind.core.plan.AssertionModeInput.defaults()),
        ExampleSteps.step(
            new StepId("step-01-ensure-sheet"),
            ExampleSelectors.sheet("Ledger"),
            new WorkbookMutationAction.EnsureSheet()),
        ExampleSteps.step(
            new StepId("step-02-append-header"),
            ExampleSelectors.sheet("Ledger"),
            new CellMutationAction.AppendRow(
                ExampleCellValues.row(
                    ExampleCellValues.text("Team"),
                    ExampleCellValues.text("Task"),
                    ExampleCellValues.text("Hours")))),
        ExampleSteps.step(
            new StepId("step-03-append-ops"),
            ExampleSelectors.sheet("Ledger"),
            new CellMutationAction.AppendRow(
                ExampleCellValues.row(
                    ExampleCellValues.text("Ops"),
                    ExampleCellValues.text("Badge prep"),
                    ExampleCellValues.number(6.5d)))),
        ExampleSteps.step(
            new StepId("step-04-append-facilities"),
            ExampleSelectors.sheet("Ledger"),
            new CellMutationAction.AppendRow(
                ExampleCellValues.row(
                    ExampleCellValues.text("Facilities"),
                    ExampleCellValues.text("Desk setup"),
                    ExampleCellValues.number(4.0d)))),
        ExampleSteps.read(
            new StepId("workbook"),
            ExampleSelectors.workbook(),
            new WorkbookIntrospectionQuery.GetWorkbookSummary()),
        ExampleSteps.read(
            new StepId("ledger-sheet"),
            ExampleSelectors.sheet("Ledger"),
            new SheetIntrospectionQuery.GetSheetSummary()));
  }

  static WorkbookPlan fileHyperlinkHealthPlan(ExamplePathLayout paths) {
    return ExampleWorkbookPlans.defaultExecutionPlan(
        new PlanId("file-hyperlink-health-workflow"),
        new WorkbookPlan.WorkbookSource.New(),
        ExampleWorkbookPlans.saveAs(
            paths.generatedWorkbook("gridgrind-file-hyperlink-health.xlsx")),
        ExampleSteps.step(
            new StepId("step-01-ensure-sheet"),
            ExampleSelectors.sheet("Links"),
            new WorkbookMutationAction.EnsureSheet()),
        ExampleSteps.step(
            new StepId("step-02-set-range"),
            ExampleSelectors.range("Links", "A1:B4"),
            new CellMutationAction.SetRange(
                ExampleCellValues.rows(
                    ExampleCellValues.row(
                        ExampleCellValues.text("Label"), ExampleCellValues.text("Destination")),
                    ExampleCellValues.row(
                        ExampleCellValues.textFile(
                            paths.asset("file-hyperlink-assets/request-label.txt")),
                        ExampleCellValues.text("support/expense policy 2026.pdf")),
                    ExampleCellValues.row(
                        ExampleCellValues.text("Absolute checklist URI"),
                        ExampleCellValues.text("file:///tmp/quarterly%20close/checklist.xlsx")),
                    ExampleCellValues.row(
                        ExampleCellValues.text("Workbook section"),
                        ExampleCellValues.text("Links!B2"))))),
        ExampleSteps.step(
            new StepId("step-03-relative-file-link"),
            ExampleSelectors.cell("Links", "A2"),
            new CellMutationAction.SetHyperlink(
                new HyperlinkTarget.File("support/expense policy 2026.pdf"))),
        ExampleSteps.step(
            new StepId("step-04-absolute-file-link"),
            ExampleSelectors.cell("Links", "A3"),
            new CellMutationAction.SetHyperlink(
                new HyperlinkTarget.File("file:///tmp/quarterly%20close/checklist.xlsx"))),
        ExampleSteps.step(
            new StepId("step-05-document-link"),
            ExampleSelectors.cell("Links", "A4"),
            new CellMutationAction.SetHyperlink(new HyperlinkTarget.Document("Links!B2"))),
        ExampleSteps.read(
            new StepId("hyperlinks"),
            ExampleSelectors.cells("Links", "A2", "A3", "A4"),
            new SheetIntrospectionQuery.GetHyperlinks()),
        ExampleSteps.read(
            new StepId("hyperlink-health"),
            ExampleSelectors.sheets("Links"),
            new InspectionAnalysisQuery.AnalyzeHyperlinkHealth()));
  }

  static WorkbookPlan introspectionAnalysisPlan(ExamplePathLayout paths) {
    return ExampleWorkbookPlans.defaultExecutionPlan(
        new PlanId("introspection-analysis-workflow"),
        new WorkbookPlan.WorkbookSource.New(),
        ExampleWorkbookPlans.saveAs(
            paths.generatedWorkbook("gridgrind-introspection-analysis.xlsx")),
        ExampleSteps.step(
            new StepId("step-01-ensure-sheet"),
            ExampleSelectors.sheet("Dashboard"),
            new WorkbookMutationAction.EnsureSheet()),
        ExampleSteps.step(
            new StepId("step-02-set-range"),
            ExampleSelectors.range("Dashboard", "A1:C4"),
            new CellMutationAction.SetRange(
                ExampleCellValues.rows(
                    ExampleCellValues.row(
                        ExampleCellValues.text("Metric"),
                        ExampleCellValues.text("Value"),
                        ExampleCellValues.text("Notes")),
                    ExampleCellValues.row(
                        ExampleCellValues.text("Revenue"),
                        ExampleCellValues.number(125000.25d),
                        ExampleCellValues.text("Closed month")),
                    ExampleCellValues.row(
                        ExampleCellValues.text("Margin"),
                        ExampleCellValues.number(0.42d),
                        ExampleCellValues.text("Target 0.40")),
                    ExampleCellValues.row(
                        ExampleCellValues.text("Forecast"),
                        ExampleCellValues.formula("B2*(1+B3)"),
                        ExampleCellValues.text("Projected next month"))))),
        ExampleSteps.step(
            new StepId("step-03-set-hyperlink"),
            ExampleSelectors.cell("Dashboard", "A1"),
            new CellMutationAction.SetHyperlink(
                new HyperlinkTarget.Url("https://example.com/dashboard-handbook"))),
        ExampleSteps.step(
            new StepId("step-04-set-comment"),
            ExampleSelectors.cell("Dashboard", "B4"),
            new CellMutationAction.SetComment(
                CommentInput.plain(
                    TextSourceInput.inline("Forecast uses the revenue and margin rows above."),
                    "GridGrind",
                    true))),
        ExampleSteps.step(
            new StepId("step-05-set-named-range"),
            new NamedRangeSelector.WorkbookScope("ForecastValue"),
            new StructuredMutationAction.SetNamedRange(
                "ForecastValue",
                new NamedRangeScope.Workbook(),
                NamedRangeTarget.range(new SheetName("Dashboard"), new RangeRef("B4")))),
        ExampleSteps.read(
            new StepId("workbook"),
            ExampleSelectors.workbook(),
            new WorkbookIntrospectionQuery.GetWorkbookSummary()),
        ExampleSteps.read(
            new StepId("formula-surface"),
            ExampleSelectors.sheets("Dashboard"),
            new InspectionSurfaceQuery.GetFormulaSurface()),
        ExampleSteps.read(
            new StepId("schema"),
            ExampleSelectors.window("Dashboard", "A1", 4, 3),
            new InspectionSurfaceQuery.GetSheetSchema()),
        ExampleSteps.read(
            new StepId("named-range-surface"),
            new NamedRangeSelector.ByName("ForecastValue"),
            new InspectionSurfaceQuery.GetNamedRangeSurface()),
        ExampleSteps.read(
            new StepId("formula-health"),
            ExampleSelectors.sheets("Dashboard"),
            new InspectionAnalysisQuery.AnalyzeFormulaHealth()),
        ExampleSteps.read(
            new StepId("hyperlink-health"),
            ExampleSelectors.sheets("Dashboard"),
            new InspectionAnalysisQuery.AnalyzeHyperlinkHealth()),
        ExampleSteps.read(
            new StepId("named-range-health"),
            new NamedRangeSelector.ByName("ForecastValue"),
            new InspectionAnalysisQuery.AnalyzeNamedRangeHealth()),
        ExampleSteps.read(
            new StepId("workbook-findings"),
            ExampleSelectors.workbook(),
            new InspectionAnalysisQuery.AnalyzeWorkbookFindings()));
  }
}
