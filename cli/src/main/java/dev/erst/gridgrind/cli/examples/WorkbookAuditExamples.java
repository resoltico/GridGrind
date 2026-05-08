package dev.erst.gridgrind.cli.examples;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.StructuredMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.dto.CalculationPolicyInput;
import dev.erst.gridgrind.contract.dto.CalculationStrategyInput;
import dev.erst.gridgrind.contract.dto.CommentInput;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.ExecutionPolicyInput;
import dev.erst.gridgrind.contract.dto.HyperlinkTarget;
import dev.erst.gridgrind.contract.dto.NamedRangeScope;
import dev.erst.gridgrind.contract.dto.NamedRangeTarget;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.InspectionAnalysisQuery;
import dev.erst.gridgrind.contract.query.InspectionSurfaceQuery;
import dev.erst.gridgrind.contract.query.SheetIntrospectionQuery;
import dev.erst.gridgrind.contract.query.WorkbookIntrospectionQuery;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.source.TextSourceInput;

/** Generated examples for inspection-heavy, health-focused, and execution-mode workflows. */
final class WorkbookAuditExamples {
  private WorkbookAuditExamples() {}

  static GridGrindShippedExamples.ShippedExample workbookHealthExample(ExamplePathLayout paths) {
    return ExamplePlanSupport.example(
        "WORKBOOK_HEALTH",
        "workbook-health-request.json",
        "Compact no-save workbook-health pass with targeted formula and aggregate findings.",
        ExamplePlanSupport.defaultExecutionPlan(
            "workbook-health-workflow",
            new WorkbookPlan.WorkbookSource.New(),
            new WorkbookPlan.WorkbookPersistence.None(),
            ExamplePlanSupport.step(
                "step-01-ensure-budget-review",
                ExamplePlanSupport.sheet("Budget Review"),
                new WorkbookMutationAction.EnsureSheet()),
            ExamplePlanSupport.step(
                "step-02-ensure-summary",
                ExamplePlanSupport.sheet("Summary"),
                new WorkbookMutationAction.EnsureSheet()),
            ExamplePlanSupport.step(
                "step-03-set-budget-header",
                ExamplePlanSupport.cell("Budget Review", "A1"),
                new CellMutationAction.SetCell(ExamplePlanSupport.text("Amount"))),
            ExamplePlanSupport.step(
                "step-04-set-budget-value",
                ExamplePlanSupport.cell("Budget Review", "B1"),
                new CellMutationAction.SetCell(ExamplePlanSupport.number(1200.0d))),
            ExamplePlanSupport.step(
                "step-05-set-summary-formula",
                ExamplePlanSupport.cell("Summary", "A1"),
                new CellMutationAction.SetCell(ExamplePlanSupport.formula("'Budget Review'!B1"))),
            ExamplePlanSupport.read(
                "summary-sheet",
                ExamplePlanSupport.sheet("Summary"),
                new SheetIntrospectionQuery.GetSheetSummary()),
            ExamplePlanSupport.read(
                "formula-health",
                ExamplePlanSupport.sheets("Summary"),
                new InspectionAnalysisQuery.AnalyzeFormulaHealth()),
            ExamplePlanSupport.read(
                "lint",
                ExamplePlanSupport.workbook(),
                new InspectionAnalysisQuery.AnalyzeWorkbookFindings()),
            ExamplePlanSupport.read(
                "summary-cells",
                ExamplePlanSupport.cells("Summary", "A1"),
                new SheetIntrospectionQuery.GetCells())));
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
                ExecutionModeInput.writeMode(ExecutionModeInput.WriteMode.STREAMING_WRITE),
                dev.erst.gridgrind.contract.dto.ExecutionJournalInput.defaults(),
                new CalculationPolicyInput(new CalculationStrategyInput.DoNotCalculate(), true)),
            ExamplePlanSupport.step(
                "step-01-ensure-sheet",
                ExamplePlanSupport.sheet("Ledger"),
                new WorkbookMutationAction.EnsureSheet()),
            ExamplePlanSupport.step(
                "step-02-append-header",
                ExamplePlanSupport.sheet("Ledger"),
                new CellMutationAction.AppendRow(
                    java.util.List.of(
                        ExamplePlanSupport.text("Team"),
                        ExamplePlanSupport.text("Task"),
                        ExamplePlanSupport.text("Hours")))),
            ExamplePlanSupport.step(
                "step-03-append-ops",
                ExamplePlanSupport.sheet("Ledger"),
                new CellMutationAction.AppendRow(
                    java.util.List.of(
                        ExamplePlanSupport.text("Ops"),
                        ExamplePlanSupport.text("Badge prep"),
                        ExamplePlanSupport.number(6.5d)))),
            ExamplePlanSupport.step(
                "step-04-append-facilities",
                ExamplePlanSupport.sheet("Ledger"),
                new CellMutationAction.AppendRow(
                    java.util.List.of(
                        ExamplePlanSupport.text("Facilities"),
                        ExamplePlanSupport.text("Desk setup"),
                        ExamplePlanSupport.number(4.0d)))),
            ExamplePlanSupport.read(
                "workbook",
                ExamplePlanSupport.workbook(),
                new WorkbookIntrospectionQuery.GetWorkbookSummary()),
            ExamplePlanSupport.read(
                "ledger-sheet",
                ExamplePlanSupport.sheet("Ledger"),
                new SheetIntrospectionQuery.GetSheetSummary())));
  }

  static GridGrindShippedExamples.ShippedExample fileHyperlinkHealthExample(
      ExamplePathLayout paths) {
    return ExamplePlanSupport.example(
        "FILE_HYPERLINK_HEALTH",
        "file-hyperlink-health-request.json",
        "File and document hyperlink authoring with explicit hyperlink-health analysis.",
        ExamplePlanSupport.defaultExecutionPlan(
            "file-hyperlink-health-workflow",
            new WorkbookPlan.WorkbookSource.New(),
            ExamplePlanSupport.saveAs(
                paths.generatedWorkbook("gridgrind-file-hyperlink-health.xlsx")),
            ExamplePlanSupport.step(
                "step-01-ensure-sheet",
                ExamplePlanSupport.sheet("Links"),
                new WorkbookMutationAction.EnsureSheet()),
            ExamplePlanSupport.step(
                "step-02-set-range",
                ExamplePlanSupport.range("Links", "A1:B4"),
                new CellMutationAction.SetRange(
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
                new CellMutationAction.SetHyperlink(
                    new HyperlinkTarget.File("support/expense policy 2026.pdf"))),
            ExamplePlanSupport.step(
                "step-04-absolute-file-link",
                ExamplePlanSupport.cell("Links", "A3"),
                new CellMutationAction.SetHyperlink(
                    new HyperlinkTarget.File("file:///tmp/quarterly%20close/checklist.xlsx"))),
            ExamplePlanSupport.step(
                "step-05-document-link",
                ExamplePlanSupport.cell("Links", "A4"),
                new CellMutationAction.SetHyperlink(new HyperlinkTarget.Document("Links!B2"))),
            ExamplePlanSupport.read(
                "hyperlinks",
                ExamplePlanSupport.cells("Links", "A2", "A3", "A4"),
                new SheetIntrospectionQuery.GetHyperlinks()),
            ExamplePlanSupport.read(
                "hyperlink-health",
                ExamplePlanSupport.sheets("Links"),
                new InspectionAnalysisQuery.AnalyzeHyperlinkHealth())));
  }

  static GridGrindShippedExamples.ShippedExample introspectionAnalysisExample(
      ExamplePathLayout paths) {
    return ExamplePlanSupport.example(
        "INTROSPECTION_ANALYSIS",
        "introspection-analysis-request.json",
        "Batch factual reads plus formula, hyperlink, named-range, and aggregate workbook analysis.",
        ExamplePlanSupport.defaultExecutionPlan(
            "introspection-analysis-workflow",
            new WorkbookPlan.WorkbookSource.New(),
            ExamplePlanSupport.saveAs(
                paths.generatedWorkbook("gridgrind-introspection-analysis.xlsx")),
            ExamplePlanSupport.step(
                "step-01-ensure-sheet",
                ExamplePlanSupport.sheet("Dashboard"),
                new WorkbookMutationAction.EnsureSheet()),
            ExamplePlanSupport.step(
                "step-02-set-range",
                ExamplePlanSupport.range("Dashboard", "A1:C4"),
                new CellMutationAction.SetRange(
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
                new CellMutationAction.SetHyperlink(
                    new HyperlinkTarget.Url("https://example.com/dashboard-handbook"))),
            ExamplePlanSupport.step(
                "step-04-set-comment",
                ExamplePlanSupport.cell("Dashboard", "B4"),
                new CellMutationAction.SetComment(
                    CommentInput.plain(
                        TextSourceInput.inline("Forecast uses the revenue and margin rows above."),
                        "GridGrind",
                        true))),
            ExamplePlanSupport.step(
                "step-05-set-named-range",
                new NamedRangeSelector.WorkbookScope("ForecastValue"),
                new StructuredMutationAction.SetNamedRange(
                    "ForecastValue",
                    new NamedRangeScope.Workbook(),
                    NamedRangeTarget.range("Dashboard", "B4"))),
            ExamplePlanSupport.read(
                "workbook",
                ExamplePlanSupport.workbook(),
                new WorkbookIntrospectionQuery.GetWorkbookSummary()),
            ExamplePlanSupport.read(
                "formula-surface",
                ExamplePlanSupport.sheets("Dashboard"),
                new InspectionSurfaceQuery.GetFormulaSurface()),
            ExamplePlanSupport.read(
                "schema",
                ExamplePlanSupport.window("Dashboard", "A1", 4, 3),
                new InspectionSurfaceQuery.GetSheetSchema()),
            ExamplePlanSupport.read(
                "named-range-surface",
                new NamedRangeSelector.ByName("ForecastValue"),
                new InspectionSurfaceQuery.GetNamedRangeSurface()),
            ExamplePlanSupport.read(
                "formula-health",
                ExamplePlanSupport.sheets("Dashboard"),
                new InspectionAnalysisQuery.AnalyzeFormulaHealth()),
            ExamplePlanSupport.read(
                "hyperlink-health",
                ExamplePlanSupport.sheets("Dashboard"),
                new InspectionAnalysisQuery.AnalyzeHyperlinkHealth()),
            ExamplePlanSupport.read(
                "named-range-health",
                new NamedRangeSelector.ByName("ForecastValue"),
                new InspectionAnalysisQuery.AnalyzeNamedRangeHealth()),
            ExamplePlanSupport.read(
                "workbook-findings",
                ExamplePlanSupport.workbook(),
                new InspectionAnalysisQuery.AnalyzeWorkbookFindings())));
  }
}
