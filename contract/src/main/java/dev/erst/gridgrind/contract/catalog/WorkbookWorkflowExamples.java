package dev.erst.gridgrind.contract.catalog;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.assertion.ExpectedCellValue;
import dev.erst.gridgrind.contract.dto.CalculationPolicyInput;
import dev.erst.gridgrind.contract.dto.CalculationStrategyInput;
import dev.erst.gridgrind.contract.dto.CellAlignmentInput;
import dev.erst.gridgrind.contract.dto.CellStyleInput;
import dev.erst.gridgrind.contract.dto.CommentInput;
import dev.erst.gridgrind.contract.dto.ExecutionJournalInput;
import dev.erst.gridgrind.contract.dto.ExecutionJournalLevel;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.HyperlinkTarget;
import dev.erst.gridgrind.contract.dto.NamedRangeScope;
import dev.erst.gridgrind.contract.dto.NamedRangeTarget;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.excel.foundation.AnalysisSeverity;
import dev.erst.gridgrind.excel.foundation.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.foundation.ExcelVerticalAlignment;
import java.util.List;

/** Workflow-centric generated examples that focus on mutation, analysis, and verification. */
final class WorkbookWorkflowExamples {
  private WorkbookWorkflowExamples() {}

  static GridGrindShippedExamples.ShippedExample budgetExample(ExamplePathLayout paths) {
    return ExamplePlanSupport.example(
        "BUDGET",
        "budget-request.json",
        "Selector-first budget sheet with styling, formula totals, readback, and schema inspection.",
        ExamplePlanSupport.plan(
            "budget-workflow",
            new WorkbookPlan.WorkbookSource.New(),
            ExamplePlanSupport.saveAs(paths.generatedWorkbook("gridgrind-budget.xlsx")),
            null,
            ExamplePlanSupport.step(
                "step-01-ensure-sheet",
                ExamplePlanSupport.sheet("Budget"),
                new MutationAction.EnsureSheet()),
            ExamplePlanSupport.step(
                "step-02-set-range",
                ExamplePlanSupport.range("Budget", "A1:C3"),
                new MutationAction.SetRange(
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
                new MutationAction.ApplyStyle(
                    new CellStyleInput(
                        null,
                        new CellAlignmentInput(
                            true,
                            ExcelHorizontalAlignment.CENTER,
                            ExcelVerticalAlignment.CENTER,
                            null,
                            null),
                        new dev.erst.gridgrind.contract.dto.CellFontInput(
                            true, null, null, null, null, null, null),
                        null,
                        null,
                        null))),
            ExamplePlanSupport.step(
                "step-04-apply-number-style",
                ExamplePlanSupport.range("Budget", "B2:B4"),
                new MutationAction.ApplyStyle(
                    new CellStyleInput(
                        "#,##0.00",
                        new CellAlignmentInput(
                            null, ExcelHorizontalAlignment.RIGHT, null, null, null),
                        null,
                        null,
                        null,
                        null))),
            ExamplePlanSupport.step(
                "step-05-set-total-label",
                ExamplePlanSupport.cell("Budget", "A4"),
                new MutationAction.SetCell(ExamplePlanSupport.text("Total"))),
            ExamplePlanSupport.step(
                "step-06-set-total-formula",
                ExamplePlanSupport.cell("Budget", "B4"),
                new MutationAction.SetCell(ExamplePlanSupport.formula("SUM(B2:B3)"))),
            ExamplePlanSupport.step(
                "step-07-auto-size",
                ExamplePlanSupport.sheet("Budget"),
                new MutationAction.AutoSizeColumns()),
            ExamplePlanSupport.read(
                "workbook",
                ExamplePlanSupport.workbook(),
                new InspectionQuery.GetWorkbookSummary()),
            ExamplePlanSupport.read(
                "cells",
                ExamplePlanSupport.cells("Budget", "A1", "B4", "C2"),
                new InspectionQuery.GetCells()),
            ExamplePlanSupport.read(
                "window",
                ExamplePlanSupport.window("Budget", "A1", 4, 3),
                new InspectionQuery.GetWindow()),
            ExamplePlanSupport.read(
                "schema",
                ExamplePlanSupport.window("Budget", "A1", 4, 3),
                new InspectionQuery.GetSheetSchema()),
            ExamplePlanSupport.read(
                "formula-surface",
                ExamplePlanSupport.sheets("Budget"),
                new InspectionQuery.GetFormulaSurface())));
  }

  static GridGrindShippedExamples.ShippedExample workbookHealthExample(ExamplePathLayout paths) {
    return ExamplePlanSupport.example(
        "WORKBOOK_HEALTH",
        "workbook-health-request.json",
        "Compact no-save workbook-health pass with targeted formula and aggregate findings.",
        ExamplePlanSupport.plan(
            "workbook-health-workflow",
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            null,
            ExamplePlanSupport.step(
                "step-01-ensure-budget-review",
                ExamplePlanSupport.sheet("Budget Review"),
                new MutationAction.EnsureSheet()),
            ExamplePlanSupport.step(
                "step-02-ensure-summary",
                ExamplePlanSupport.sheet("Summary"),
                new MutationAction.EnsureSheet()),
            ExamplePlanSupport.step(
                "step-03-set-budget-header",
                ExamplePlanSupport.cell("Budget Review", "A1"),
                new MutationAction.SetCell(ExamplePlanSupport.text("Amount"))),
            ExamplePlanSupport.step(
                "step-04-set-budget-value",
                ExamplePlanSupport.cell("Budget Review", "B1"),
                new MutationAction.SetCell(ExamplePlanSupport.number(1200.0d))),
            ExamplePlanSupport.step(
                "step-05-set-summary-formula",
                ExamplePlanSupport.cell("Summary", "A1"),
                new MutationAction.SetCell(ExamplePlanSupport.formula("'Budget Review'!B1"))),
            ExamplePlanSupport.read(
                "summary-sheet",
                ExamplePlanSupport.sheet("Summary"),
                new InspectionQuery.GetSheetSummary()),
            ExamplePlanSupport.read(
                "formula-health",
                ExamplePlanSupport.sheets("Summary"),
                new InspectionQuery.AnalyzeFormulaHealth()),
            ExamplePlanSupport.read(
                "lint",
                ExamplePlanSupport.workbook(),
                new InspectionQuery.AnalyzeWorkbookFindings()),
            ExamplePlanSupport.read(
                "summary-cells",
                ExamplePlanSupport.cells("Summary", "A1"),
                new InspectionQuery.GetCells())));
  }

  static GridGrindShippedExamples.ShippedExample sheetMaintenanceExample(ExamplePathLayout paths) {
    return ExamplePlanSupport.example(
        "SHEET_MAINTENANCE",
        "sheet-maintenance-request.json",
        "Copy-sheet maintenance walkthrough with comment reread and workbook findings.",
        ExamplePlanSupport.plan(
            "sheet-maintenance-workflow",
            new WorkbookPlan.WorkbookSource.New(),
            ExamplePlanSupport.saveAs(paths.generatedWorkbook("gridgrind-sheet-maintenance.xlsx")),
            null,
            ExamplePlanSupport.step(
                "step-01-ensure-template",
                ExamplePlanSupport.sheet("Template"),
                new MutationAction.EnsureSheet()),
            ExamplePlanSupport.step(
                "step-02-set-range",
                ExamplePlanSupport.range("Template", "A1:B3"),
                new MutationAction.SetRange(
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
                new MutationAction.SetComment(
                    new CommentInput(
                        TextSourceInput.inline("Template owner column"), "GridGrind", false))),
            ExamplePlanSupport.step(
                "step-04-copy-sheet",
                ExamplePlanSupport.sheet("Template"),
                new MutationAction.CopySheet("Template Copy", null)),
            ExamplePlanSupport.read(
                "step-05-read-comments",
                new CellSelector.AllUsedInSheet("Template Copy"),
                new InspectionQuery.GetComments()),
            ExamplePlanSupport.read(
                "step-06-read-workbook-findings",
                ExamplePlanSupport.workbook(),
                new InspectionQuery.AnalyzeWorkbookFindings())));
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
            new ExecutionPolicyInput(
                null, new ExecutionJournalInput(ExecutionJournalLevel.VERBOSE), null),
            ExamplePlanSupport.step(
                "ensure-budget",
                ExamplePlanSupport.sheet("Budget"),
                new MutationAction.EnsureSheet()),
            ExamplePlanSupport.step(
                "set-title",
                ExamplePlanSupport.cell("Budget", "A1"),
                new MutationAction.SetCell(ExamplePlanSupport.text("Quarterly Budget"))),
            ExamplePlanSupport.step(
                "set-total",
                ExamplePlanSupport.cell("Budget", "B2"),
                new MutationAction.SetCell(ExamplePlanSupport.number(1200.0d))),
            ExamplePlanSupport.assertStep(
                "assert-title",
                ExamplePlanSupport.cell("Budget", "A1"),
                new Assertion.CellValue(new ExpectedCellValue.Text("Quarterly Budget"))),
            ExamplePlanSupport.assertStep(
                "assert-total",
                ExamplePlanSupport.cell("Budget", "B2"),
                new Assertion.CellValue(new ExpectedCellValue.NumericValue(1200.0d))),
            ExamplePlanSupport.assertStep(
                "assert-formula-health",
                ExamplePlanSupport.sheet("Budget"),
                new Assertion.AnalysisMaxSeverity(
                    new InspectionQuery.AnalyzeFormulaHealth(), AnalysisSeverity.INFO)),
            ExamplePlanSupport.read(
                "read-budget",
                ExamplePlanSupport.cells("Budget", "A1", "B2"),
                new InspectionQuery.GetCells())));
  }

  static GridGrindShippedExamples.ShippedExample largeFileModesExample(ExamplePathLayout paths) {
    return ExamplePlanSupport.example(
        "LARGE_FILE_MODES",
        "large-file-modes-request.json",
        "Low-memory STREAMING_WRITE plan with append-only rows and recalc-on-open flagging.",
        ExamplePlanSupport.plan(
            "large-file-modes-workflow",
            new WorkbookPlan.WorkbookSource.New(),
            ExamplePlanSupport.saveAs(paths.generatedWorkbook("gridgrind-large-file-modes.xlsx")),
            new ExecutionPolicyInput(
                new ExecutionModeInput(null, ExecutionModeInput.WriteMode.STREAMING_WRITE),
                null,
                new CalculationPolicyInput(new CalculationStrategyInput.DoNotCalculate(), true)),
            ExamplePlanSupport.step(
                "step-01-ensure-sheet",
                ExamplePlanSupport.sheet("Ledger"),
                new MutationAction.EnsureSheet()),
            ExamplePlanSupport.step(
                "step-02-append-header",
                ExamplePlanSupport.sheet("Ledger"),
                new MutationAction.AppendRow(
                    List.of(
                        ExamplePlanSupport.text("Team"),
                        ExamplePlanSupport.text("Task"),
                        ExamplePlanSupport.text("Hours")))),
            ExamplePlanSupport.step(
                "step-03-append-ops",
                ExamplePlanSupport.sheet("Ledger"),
                new MutationAction.AppendRow(
                    List.of(
                        ExamplePlanSupport.text("Ops"),
                        ExamplePlanSupport.text("Badge prep"),
                        ExamplePlanSupport.number(6.5d)))),
            ExamplePlanSupport.step(
                "step-04-append-facilities",
                ExamplePlanSupport.sheet("Ledger"),
                new MutationAction.AppendRow(
                    List.of(
                        ExamplePlanSupport.text("Facilities"),
                        ExamplePlanSupport.text("Desk setup"),
                        ExamplePlanSupport.number(4.0d)))),
            ExamplePlanSupport.read(
                "workbook",
                ExamplePlanSupport.workbook(),
                new InspectionQuery.GetWorkbookSummary()),
            ExamplePlanSupport.read(
                "ledger-sheet",
                ExamplePlanSupport.sheet("Ledger"),
                new InspectionQuery.GetSheetSummary())));
  }

  static GridGrindShippedExamples.ShippedExample fileHyperlinkHealthExample(
      ExamplePathLayout paths) {
    return ExamplePlanSupport.example(
        "FILE_HYPERLINK_HEALTH",
        "file-hyperlink-health-request.json",
        "File and document hyperlink authoring with explicit hyperlink-health analysis.",
        ExamplePlanSupport.plan(
            "file-hyperlink-health-workflow",
            new WorkbookPlan.WorkbookSource.New(),
            ExamplePlanSupport.saveAs(
                paths.generatedWorkbook("gridgrind-file-hyperlink-health.xlsx")),
            null,
            ExamplePlanSupport.step(
                "step-01-ensure-sheet",
                ExamplePlanSupport.sheet("Links"),
                new MutationAction.EnsureSheet()),
            ExamplePlanSupport.step(
                "step-02-set-range",
                ExamplePlanSupport.range("Links", "A1:B4"),
                new MutationAction.SetRange(
                    ExamplePlanSupport.rows(
                        ExamplePlanSupport.row(
                            ExamplePlanSupport.text("Label"),
                            ExamplePlanSupport.text("Destination")),
                        ExamplePlanSupport.row(
                            ExamplePlanSupport.text("Relative policy PDF"),
                            ExamplePlanSupport.text("support/expense policy 2026.pdf")),
                        ExamplePlanSupport.row(
                            ExamplePlanSupport.text("Absolute checklist URI"),
                            ExamplePlanSupport.text(
                                "file:///tmp/quarterly%20close/checklist.xlsx")),
                        ExamplePlanSupport.row(
                            ExamplePlanSupport.text("Workbook section"),
                            ExamplePlanSupport.text("Links!B2"))))),
            ExamplePlanSupport.step(
                "step-03-relative-file-link",
                ExamplePlanSupport.cell("Links", "A2"),
                new MutationAction.SetHyperlink(
                    new HyperlinkTarget.File("support/expense policy 2026.pdf"))),
            ExamplePlanSupport.step(
                "step-04-absolute-file-link",
                ExamplePlanSupport.cell("Links", "A3"),
                new MutationAction.SetHyperlink(
                    new HyperlinkTarget.File("file:///tmp/quarterly%20close/checklist.xlsx"))),
            ExamplePlanSupport.step(
                "step-05-document-link",
                ExamplePlanSupport.cell("Links", "A4"),
                new MutationAction.SetHyperlink(new HyperlinkTarget.Document("Links!B2"))),
            ExamplePlanSupport.read(
                "hyperlinks",
                ExamplePlanSupport.cells("Links", "A2", "A3", "A4"),
                new InspectionQuery.GetHyperlinks()),
            ExamplePlanSupport.read(
                "hyperlink-health",
                ExamplePlanSupport.sheets("Links"),
                new InspectionQuery.AnalyzeHyperlinkHealth())));
  }

  static GridGrindShippedExamples.ShippedExample introspectionAnalysisExample(
      ExamplePathLayout paths) {
    return ExamplePlanSupport.example(
        "INTROSPECTION_ANALYSIS",
        "introspection-analysis-request.json",
        "Batch factual reads plus formula, hyperlink, named-range, and aggregate workbook analysis.",
        ExamplePlanSupport.plan(
            "introspection-analysis-workflow",
            new WorkbookPlan.WorkbookSource.New(),
            ExamplePlanSupport.saveAs(
                paths.generatedWorkbook("gridgrind-introspection-analysis.xlsx")),
            null,
            ExamplePlanSupport.step(
                "step-01-ensure-sheet",
                ExamplePlanSupport.sheet("Dashboard"),
                new MutationAction.EnsureSheet()),
            ExamplePlanSupport.step(
                "step-02-set-range",
                ExamplePlanSupport.range("Dashboard", "A1:C4"),
                new MutationAction.SetRange(
                    ExamplePlanSupport.rows(
                        ExamplePlanSupport.row(
                            ExamplePlanSupport.text("Metric"),
                            ExamplePlanSupport.text("Value"),
                            ExamplePlanSupport.text("Notes")),
                        ExamplePlanSupport.row(
                            ExamplePlanSupport.text("Revenue"),
                            ExamplePlanSupport.number(125000.25d),
                            ExamplePlanSupport.text("Closed month")),
                        ExamplePlanSupport.row(
                            ExamplePlanSupport.text("Margin"),
                            ExamplePlanSupport.number(0.42d),
                            ExamplePlanSupport.text("Target 0.40")),
                        ExamplePlanSupport.row(
                            ExamplePlanSupport.text("Forecast"),
                            ExamplePlanSupport.formula("B2*(1+B3)"),
                            ExamplePlanSupport.text("Projected next month"))))),
            ExamplePlanSupport.step(
                "step-03-set-hyperlink",
                ExamplePlanSupport.cell("Dashboard", "A1"),
                new MutationAction.SetHyperlink(
                    new HyperlinkTarget.Url("https://example.com/dashboard-handbook"))),
            ExamplePlanSupport.step(
                "step-04-set-comment",
                ExamplePlanSupport.cell("Dashboard", "B4"),
                new MutationAction.SetComment(
                    new CommentInput(
                        TextSourceInput.inline("Forecast uses the revenue and margin rows above."),
                        "GridGrind",
                        true))),
            ExamplePlanSupport.step(
                "step-05-set-named-range",
                new NamedRangeSelector.WorkbookScope("ForecastValue"),
                new MutationAction.SetNamedRange(
                    "ForecastValue",
                    new NamedRangeScope.Workbook(),
                    new NamedRangeTarget("Dashboard", "B4"))),
            ExamplePlanSupport.read(
                "workbook",
                ExamplePlanSupport.workbook(),
                new InspectionQuery.GetWorkbookSummary()),
            ExamplePlanSupport.read(
                "formula-surface",
                ExamplePlanSupport.sheets("Dashboard"),
                new InspectionQuery.GetFormulaSurface()),
            ExamplePlanSupport.read(
                "schema",
                ExamplePlanSupport.window("Dashboard", "A1", 4, 3),
                new InspectionQuery.GetSheetSchema()),
            ExamplePlanSupport.read(
                "named-range-surface",
                new NamedRangeSelector.ByName("ForecastValue"),
                new InspectionQuery.GetNamedRangeSurface()),
            ExamplePlanSupport.read(
                "formula-health",
                ExamplePlanSupport.sheets("Dashboard"),
                new InspectionQuery.AnalyzeFormulaHealth()),
            ExamplePlanSupport.read(
                "hyperlink-health",
                ExamplePlanSupport.sheets("Dashboard"),
                new InspectionQuery.AnalyzeHyperlinkHealth()),
            ExamplePlanSupport.read(
                "named-range-health",
                new NamedRangeSelector.ByName("ForecastValue"),
                new InspectionQuery.AnalyzeNamedRangeHealth()),
            ExamplePlanSupport.read(
                "workbook-findings",
                ExamplePlanSupport.workbook(),
                new InspectionQuery.AnalyzeWorkbookFindings())));
  }
}
