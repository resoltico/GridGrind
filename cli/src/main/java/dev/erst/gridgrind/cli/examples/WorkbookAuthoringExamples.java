package dev.erst.gridgrind.cli.examples;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.assertion.AnalysisAssertion;
import dev.erst.gridgrind.contract.assertion.CellAssertion;
import dev.erst.gridgrind.contract.assertion.ExpectedCellValue;
import dev.erst.gridgrind.contract.dto.ArrayFormulaInput;
import dev.erst.gridgrind.contract.dto.CellAlignmentInput;
import dev.erst.gridgrind.contract.dto.CellStyleInput;
import dev.erst.gridgrind.contract.dto.CommentInput;
import dev.erst.gridgrind.contract.dto.ExecutionJournalInput;
import dev.erst.gridgrind.contract.dto.ExecutionJournalLevel;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.InspectionAnalysisQuery;
import dev.erst.gridgrind.contract.query.InspectionSurfaceQuery;
import dev.erst.gridgrind.contract.query.SheetIntrospectionQuery;
import dev.erst.gridgrind.contract.query.WorkbookIntrospectionQuery;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.excel.foundation.AnalysisSeverity;
import dev.erst.gridgrind.excel.foundation.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.foundation.ExcelVerticalAlignment;
import java.util.Optional;

/** Generated examples for core workbook authoring, maintenance, assertions, and array formulas. */
final class WorkbookAuthoringExamples {
  private WorkbookAuthoringExamples() {}

  static GridGrindShippedExamples.ShippedExample budgetExample(ExamplePathLayout paths) {
    return ExamplePlanSupport.example(
        "BUDGET",
        "budget-request.json",
        "Selector-first budget sheet with styling, formula totals, readback, and schema inspection.",
        ExamplePlanSupport.defaultExecutionPlan(
            "budget-workflow",
            new WorkbookPlan.WorkbookSource.New(),
            ExamplePlanSupport.saveAs(paths.generatedWorkbook("gridgrind-budget.xlsx")),
            ExamplePlanSupport.step(
                "step-01-ensure-sheet",
                ExamplePlanSupport.sheet("Budget"),
                new WorkbookMutationAction.EnsureSheet()),
            ExamplePlanSupport.step(
                "step-02-set-range",
                ExamplePlanSupport.range("Budget", "A1:C3"),
                new CellMutationAction.SetRange(
                    ExamplePlanSupport.rows(
                        ExamplePlanSupport.row(
                            ExamplePlanSupport.text("Item"),
                            ExamplePlanSupport.text("Amount"),
                            ExamplePlanSupport.text("Billable")),
                        ExamplePlanSupport.row(
                            ExamplePlanSupport.text("Hosting"),
                            ExamplePlanSupport.number(49.0d),
                            ExamplePlanSupport.bool(true)),
                        ExamplePlanSupport.row(
                            ExamplePlanSupport.text("Domain"),
                            ExamplePlanSupport.number(12.0d),
                            ExamplePlanSupport.bool(false))))),
            ExamplePlanSupport.step(
                "step-03-apply-header-style",
                ExamplePlanSupport.range("Budget", "A1:C1"),
                new CellMutationAction.ApplyStyle(
                    new CellStyleInput(
                        Optional.empty(),
                        Optional.of(
                            new CellAlignmentInput(
                                Optional.of(true),
                                Optional.of(ExcelHorizontalAlignment.CENTER),
                                Optional.of(ExcelVerticalAlignment.CENTER),
                                Optional.empty(),
                                Optional.empty())),
                        Optional.of(
                            new dev.erst.gridgrind.contract.dto.CellFontInput(
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
            ExamplePlanSupport.step(
                "step-04-apply-number-style",
                ExamplePlanSupport.range("Budget", "B2:B4"),
                new CellMutationAction.ApplyStyle(
                    new CellStyleInput(
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
            ExamplePlanSupport.step(
                "step-05-set-total-label",
                ExamplePlanSupport.cell("Budget", "A4"),
                new CellMutationAction.SetCell(ExamplePlanSupport.text("Total"))),
            ExamplePlanSupport.step(
                "step-06-set-total-formula",
                ExamplePlanSupport.cell("Budget", "B4"),
                new CellMutationAction.SetCell(ExamplePlanSupport.formula("SUM(B2:B3)"))),
            ExamplePlanSupport.step(
                "step-07-auto-size",
                ExamplePlanSupport.sheet("Budget"),
                new WorkbookMutationAction.AutoSizeColumns()),
            ExamplePlanSupport.read(
                "workbook",
                ExamplePlanSupport.workbook(),
                new WorkbookIntrospectionQuery.GetWorkbookSummary()),
            ExamplePlanSupport.read(
                "cells",
                ExamplePlanSupport.cells("Budget", "A1", "B4", "C2"),
                new SheetIntrospectionQuery.GetCells()),
            ExamplePlanSupport.read(
                "window",
                ExamplePlanSupport.window("Budget", "A1", 4, 3),
                new SheetIntrospectionQuery.GetWindow()),
            ExamplePlanSupport.read(
                "schema",
                ExamplePlanSupport.window("Budget", "A1", 4, 3),
                new InspectionSurfaceQuery.GetSheetSchema()),
            ExamplePlanSupport.read(
                "formula-surface",
                ExamplePlanSupport.sheets("Budget"),
                new InspectionSurfaceQuery.GetFormulaSurface())));
  }

  static GridGrindShippedExamples.ShippedExample sheetMaintenanceExample(ExamplePathLayout paths) {
    return ExamplePlanSupport.example(
        "SHEET_MAINTENANCE",
        "sheet-maintenance-request.json",
        "Copy-sheet maintenance walkthrough with comment reread and workbook findings.",
        ExamplePlanSupport.defaultExecutionPlan(
            "sheet-maintenance-workflow",
            new WorkbookPlan.WorkbookSource.New(),
            ExamplePlanSupport.saveAs(paths.generatedWorkbook("gridgrind-sheet-maintenance.xlsx")),
            ExamplePlanSupport.step(
                "step-01-ensure-template",
                ExamplePlanSupport.sheet("Template"),
                new WorkbookMutationAction.EnsureSheet()),
            ExamplePlanSupport.step(
                "step-02-set-range",
                ExamplePlanSupport.range("Template", "A1:B3"),
                new CellMutationAction.SetRange(
                    ExamplePlanSupport.rows(
                        ExamplePlanSupport.row(
                            ExamplePlanSupport.text("Owner"), ExamplePlanSupport.text("Status")),
                        ExamplePlanSupport.row(
                            ExamplePlanSupport.text("Ada"), ExamplePlanSupport.text("Ready")),
                        ExamplePlanSupport.row(
                            ExamplePlanSupport.text("Lin"), ExamplePlanSupport.text("Review"))))),
            ExamplePlanSupport.step(
                "step-03-set-comment",
                ExamplePlanSupport.cell("Template", "A1"),
                new CellMutationAction.SetComment(
                    CommentInput.plain(
                        TextSourceInput.inline("Template owner column"), "GridGrind", false))),
            ExamplePlanSupport.step(
                "step-04-copy-sheet",
                ExamplePlanSupport.sheet("Template"),
                new WorkbookMutationAction.CopySheet("Template Copy")),
            ExamplePlanSupport.read(
                "step-05-read-comments",
                new CellSelector.AllUsedInSheet("Template Copy"),
                new SheetIntrospectionQuery.GetComments()),
            ExamplePlanSupport.read(
                "step-06-read-workbook-findings",
                ExamplePlanSupport.workbook(),
                new InspectionAnalysisQuery.AnalyzeWorkbookFindings())));
  }

  static GridGrindShippedExamples.ShippedExample assertionExample(ExamplePathLayout paths) {
    return ExamplePlanSupport.example(
        "ASSERTION",
        "assertion-request.json",
        "Mutate then verify with first-class assertions, verbose journaling, and factual readback.",
        ExamplePlanSupport.plan(
            "assertion-workflow",
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExecutionPolicyInput.journal(new ExecutionJournalInput(ExecutionJournalLevel.VERBOSE)),
            ExamplePlanSupport.step(
                "ensure-budget",
                ExamplePlanSupport.sheet("Budget"),
                new WorkbookMutationAction.EnsureSheet()),
            ExamplePlanSupport.step(
                "set-title",
                ExamplePlanSupport.cell("Budget", "A1"),
                new CellMutationAction.SetCell(ExamplePlanSupport.text("Quarterly Budget"))),
            ExamplePlanSupport.step(
                "set-total",
                ExamplePlanSupport.cell("Budget", "B2"),
                new CellMutationAction.SetCell(ExamplePlanSupport.number(1200.0d))),
            ExamplePlanSupport.assertStep(
                "assert-title",
                ExamplePlanSupport.cell("Budget", "A1"),
                new CellAssertion.CellValue(new ExpectedCellValue.Text("Quarterly Budget"))),
            ExamplePlanSupport.assertStep(
                "assert-total",
                ExamplePlanSupport.cell("Budget", "B2"),
                new CellAssertion.CellValue(new ExpectedCellValue.NumericValue(1200.0d))),
            ExamplePlanSupport.assertStep(
                "assert-formula-health",
                ExamplePlanSupport.sheet("Budget"),
                new AnalysisAssertion.AnalysisMaxSeverity(
                    new InspectionAnalysisQuery.AnalyzeFormulaHealth(), AnalysisSeverity.INFO)),
            ExamplePlanSupport.read(
                "read-budget",
                ExamplePlanSupport.cells("Budget", "A1", "B2"),
                new SheetIntrospectionQuery.GetCells())));
  }

  static GridGrindShippedExamples.ShippedExample arrayFormulaExample(ExamplePathLayout paths) {
    return ExamplePlanSupport.example(
        "ARRAY_FORMULA",
        "array-formula-request.json",
        "Array-formula authoring with factual group readback and group clearing.",
        ExamplePlanSupport.defaultExecutionPlan(
            "array-formula-workflow",
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExamplePlanSupport.step(
                "ensure-calc-sheet",
                ExamplePlanSupport.sheet("Calc"),
                new WorkbookMutationAction.EnsureSheet()),
            ExamplePlanSupport.step(
                "seed-source-data",
                ExamplePlanSupport.range("Calc", "A1:C4"),
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
                            ExamplePlanSupport.number(16.0d)),
                        ExamplePlanSupport.row(
                            ExamplePlanSupport.text("Mar"),
                            ExamplePlanSupport.number(15.0d),
                            ExamplePlanSupport.number(21.0d))))),
            ExamplePlanSupport.step(
                "author-array-group",
                ExamplePlanSupport.range("Calc", "D2:D4"),
                new CellMutationAction.SetArrayFormula(
                    new ArrayFormulaInput(TextSourceInput.inline("{=B2:B4*C2:C4}")))),
            ExamplePlanSupport.read(
                "read-array-groups",
                ExamplePlanSupport.sheet("Calc"),
                new SheetIntrospectionQuery.GetArrayFormulas()),
            ExamplePlanSupport.step(
                "clear-array-group",
                ExamplePlanSupport.cell("Calc", "D3"),
                new CellMutationAction.ClearArrayFormula()),
            ExamplePlanSupport.read(
                "read-array-groups-after-clear",
                ExamplePlanSupport.sheet("Calc"),
                new SheetIntrospectionQuery.GetArrayFormulas())));
  }
}
