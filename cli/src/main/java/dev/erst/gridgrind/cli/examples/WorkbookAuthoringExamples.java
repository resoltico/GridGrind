package dev.erst.gridgrind.cli.examples;

import dev.erst.gridgrind.core.model.drawing.CommentInput;
import dev.erst.gridgrind.core.model.formula.ArrayFormulaInput;
import dev.erst.gridgrind.core.model.style.CellAlignmentInput;
import dev.erst.gridgrind.core.model.style.CellStylePatchInput;
import dev.erst.gridgrind.core.plan.AnalysisAssertion;
import dev.erst.gridgrind.core.plan.CellAssertion;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.ExecutionJournalInput;
import dev.erst.gridgrind.core.plan.ExecutionPolicyInput;
import dev.erst.gridgrind.core.plan.InspectionAnalysisQuery;
import dev.erst.gridgrind.core.plan.InspectionSurfaceQuery;
import dev.erst.gridgrind.core.plan.SheetIntrospectionQuery;
import dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.result.ExecutionJournalLevel;
import dev.erst.gridgrind.core.source.TextSourceInput;
import dev.erst.gridgrind.core.vocabulary.AnalysisSeverity;
import dev.erst.gridgrind.core.vocabulary.ExcelHorizontalAlignment;
import dev.erst.gridgrind.core.vocabulary.ExcelVerticalAlignment;
import dev.erst.gridgrind.core.vocabulary.PlanId;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import java.util.Optional;

/** Generated examples for core workbook authoring, maintenance, assertions, and array formulas. */
final class WorkbookAuthoringExamples {
  private WorkbookAuthoringExamples() {}

  static WorkbookPlan budgetPlan(ExamplePathLayout paths) {
    return ExampleWorkbookPlans.defaultExecutionPlan(
        new PlanId("budget-workflow"),
        new WorkbookPlan.WorkbookSource.New(),
        ExampleWorkbookPlans.saveAs(paths.generatedWorkbook("gridgrind-budget.xlsx")),
        ExampleSteps.step(
            new StepId("step-01-ensure-sheet"),
            ExampleSelectors.sheet("Budget"),
            new WorkbookMutationAction.EnsureSheet()),
        ExampleSteps.step(
            new StepId("step-02-set-range"),
            ExampleSelectors.range("Budget", "A1:C3"),
            new CellMutationAction.SetRange(
                ExampleCellValues.rows(
                    ExampleCellValues.row(
                        ExampleCellValues.text("Item"),
                        ExampleCellValues.text("Amount"),
                        ExampleCellValues.text("Billable")),
                    ExampleCellValues.row(
                        ExampleCellValues.text("Hosting"),
                        ExampleCellValues.number(49.0d),
                        ExampleCellValues.bool(true)),
                    ExampleCellValues.row(
                        ExampleCellValues.text("Domain"),
                        ExampleCellValues.number(12.0d),
                        ExampleCellValues.bool(false))))),
        ExampleSteps.step(
            new StepId("step-03-apply-header-style"),
            ExampleSelectors.range("Budget", "A1:C1"),
            new CellMutationAction.ApplyStyle(
                new CellStylePatchInput(
                    Optional.empty(),
                    Optional.of(
                        new CellAlignmentInput(
                            Optional.of(true),
                            Optional.of(ExcelHorizontalAlignment.CENTER),
                            Optional.of(ExcelVerticalAlignment.CENTER),
                            Optional.empty(),
                            Optional.empty())),
                    Optional.of(
                        new dev.erst.gridgrind.core.model.style.CellFontInput(
                            Optional.of(true),
                            Optional.empty(),
                            Optional.empty(),
                            Optional.empty(),
                            Optional.empty(),
                            Optional.empty(),
                            Optional.empty())),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty()))),
        ExampleSteps.step(
            new StepId("step-04-apply-number-style"),
            ExampleSelectors.range("Budget", "B2:B4"),
            new CellMutationAction.ApplyStyle(
                new CellStylePatchInput(
                    Optional.of("#,##0.00"),
                    Optional.of(
                        new CellAlignmentInput(
                            Optional.empty(),
                            Optional.of(ExcelHorizontalAlignment.RIGHT),
                            Optional.empty(),
                            Optional.empty(),
                            Optional.empty())),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty(),
                    Optional.empty()))),
        ExampleSteps.step(
            new StepId("step-05-set-total-label"),
            ExampleSelectors.cell("Budget", "A4"),
            new CellMutationAction.SetCell(ExampleCellValues.text("Total"))),
        ExampleSteps.step(
            new StepId("step-06-set-total-formula"),
            ExampleSelectors.cell("Budget", "B4"),
            new CellMutationAction.SetCell(ExampleCellValues.formula("SUM(B2:B3)"))),
        ExampleSteps.step(
            new StepId("step-07-auto-size"),
            ExampleSelectors.sheet("Budget"),
            new WorkbookMutationAction.AutoSizeColumns()),
        ExampleSteps.read(
            new StepId("workbook"),
            ExampleSelectors.workbook(),
            new WorkbookIntrospectionQuery.GetWorkbookSummary()),
        ExampleSteps.read(
            new StepId("cells"),
            ExampleSelectors.cells("Budget", "A1", "B4", "C2"),
            new SheetIntrospectionQuery.GetCells()),
        ExampleSteps.read(
            new StepId("window"),
            ExampleSelectors.window("Budget", "A1", 4, 3),
            new SheetIntrospectionQuery.GetWindow()),
        ExampleSteps.read(
            new StepId("schema"),
            ExampleSelectors.window("Budget", "A1", 4, 3),
            new InspectionSurfaceQuery.GetSheetSchema()),
        ExampleSteps.read(
            new StepId("formula-surface"),
            ExampleSelectors.sheets("Budget"),
            new InspectionSurfaceQuery.GetFormulaSurface()));
  }

  static WorkbookPlan sheetMaintenancePlan(ExamplePathLayout paths) {
    return ExampleWorkbookPlans.defaultExecutionPlan(
        new PlanId("sheet-maintenance-workflow"),
        new WorkbookPlan.WorkbookSource.New(),
        ExampleWorkbookPlans.saveAs(paths.generatedWorkbook("gridgrind-sheet-maintenance.xlsx")),
        ExampleSteps.step(
            new StepId("step-01-ensure-template"),
            ExampleSelectors.sheet("Template"),
            new WorkbookMutationAction.EnsureSheet()),
        ExampleSteps.step(
            new StepId("step-02-set-range"),
            ExampleSelectors.range("Template", "A1:B3"),
            new CellMutationAction.SetRange(
                ExampleCellValues.rows(
                    ExampleCellValues.row(
                        ExampleCellValues.text("Owner"), ExampleCellValues.text("Status")),
                    ExampleCellValues.row(
                        ExampleCellValues.text("Ada"), ExampleCellValues.text("Ready")),
                    ExampleCellValues.row(
                        ExampleCellValues.text("Lin"), ExampleCellValues.text("Review"))))),
        ExampleSteps.step(
            new StepId("step-03-set-comment"),
            ExampleSelectors.cell("Template", "A1"),
            new CellMutationAction.SetComment(
                CommentInput.plain(
                    TextSourceInput.inline("Template owner column"), "GridGrind", false))),
        ExampleSteps.step(
            new StepId("step-04-copy-sheet"),
            ExampleSelectors.sheet("Template"),
            new WorkbookMutationAction.CopySheet(new SheetName("Template Copy"))),
        ExampleSteps.read(
            new StepId("step-05-read-comments"),
            new CellSelector.AllUsedInSheet(new SheetName("Template Copy")),
            new SheetIntrospectionQuery.GetComments()),
        ExampleSteps.read(
            new StepId("step-06-read-workbook-findings"),
            ExampleSelectors.workbook(),
            new InspectionAnalysisQuery.AnalyzeWorkbookFindings()));
  }

  static WorkbookPlan assertionPlan(ExamplePathLayout paths) {
    return ExampleWorkbookPlans.plan(
        new PlanId("assertion-workflow"),
        new WorkbookPlan.WorkbookSource.New(),
        new WorkbookPlan.WorkbookPersistence.None(),
        ExecutionPolicyInput.journal(new ExecutionJournalInput(ExecutionJournalLevel.VERBOSE)),
        ExampleSteps.step(
            new StepId("ensure-budget"),
            ExampleSelectors.sheet("Budget"),
            new WorkbookMutationAction.EnsureSheet()),
        ExampleSteps.step(
            new StepId("set-title"),
            ExampleSelectors.cell("Budget", "A1"),
            new CellMutationAction.SetCell(ExampleCellValues.text("Quarterly Budget"))),
        ExampleSteps.step(
            new StepId("set-total"),
            ExampleSelectors.cell("Budget", "B2"),
            new CellMutationAction.SetCell(ExampleCellValues.number(1200.0d))),
        ExampleSteps.assertStep(
            new StepId("assert-title"),
            ExampleSelectors.cell("Budget", "A1"),
            new CellAssertion.CellValue(
                new dev.erst.gridgrind.core.model.value.CellScalarValue.Text("Quarterly Budget"))),
        ExampleSteps.assertStep(
            new StepId("assert-total"),
            ExampleSelectors.cell("Budget", "B2"),
            new CellAssertion.CellValue(
                new dev.erst.gridgrind.core.model.value.CellScalarValue.NumberValue(1200.0d))),
        ExampleSteps.assertStep(
            new StepId("assert-formula-health"),
            ExampleSelectors.sheet("Budget"),
            new AnalysisAssertion.AnalysisMaxSeverity(
                new InspectionAnalysisQuery.AnalyzeFormulaHealth(), AnalysisSeverity.INFO)),
        ExampleSteps.read(
            new StepId("read-budget"),
            ExampleSelectors.cells("Budget", "A1", "B2"),
            new SheetIntrospectionQuery.GetCells()));
  }

  static WorkbookPlan arrayFormulaPlan(ExamplePathLayout paths) {
    return ExampleWorkbookPlans.defaultExecutionPlan(
        new PlanId("array-formula-workflow"),
        new WorkbookPlan.WorkbookSource.New(),
        new WorkbookPlan.WorkbookPersistence.None(),
        ExampleSteps.step(
            new StepId("ensure-calc-sheet"),
            ExampleSelectors.sheet("Calc"),
            new WorkbookMutationAction.EnsureSheet()),
        ExampleSteps.step(
            new StepId("seed-source-data"),
            ExampleSelectors.range("Calc", "A1:C4"),
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
                        ExampleCellValues.number(16.0d)),
                    ExampleCellValues.row(
                        ExampleCellValues.text("Mar"),
                        ExampleCellValues.number(15.0d),
                        ExampleCellValues.number(21.0d))))),
        ExampleSteps.step(
            new StepId("author-array-group"),
            ExampleSelectors.range("Calc", "D2:D4"),
            new CellMutationAction.SetArrayFormula(
                new ArrayFormulaInput(TextSourceInput.inline("{=B2:B4*C2:C4}")))),
        ExampleSteps.read(
            new StepId("read-array-groups"),
            ExampleSelectors.sheet("Calc"),
            new SheetIntrospectionQuery.GetArrayFormulas()),
        ExampleSteps.step(
            new StepId("clear-array-group"),
            ExampleSelectors.cell("Calc", "D3"),
            new CellMutationAction.ClearArrayFormula()),
        ExampleSteps.read(
            new StepId("read-array-groups-after-clear"),
            ExampleSelectors.sheet("Calc"),
            new SheetIntrospectionQuery.GetArrayFormulas()));
  }
}
