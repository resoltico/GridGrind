package dev.erst.gridgrind.engine.backend.xssf;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.table.AutofilterHealthReport;
import dev.erst.gridgrind.core.model.table.ResolvedTable;
import dev.erst.gridgrind.core.model.table.TableHealthReport;
import dev.erst.gridgrind.core.model.table.TableStyleSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationHealthReport;
import dev.erst.gridgrind.core.model.validation.DataValidationRuleSpec;
import dev.erst.gridgrind.core.model.validation.ResolvedDataValidation;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.core.result.AnalysisFindingReport;
import dev.erst.gridgrind.core.result.AnalysisLocationReport;
import dev.erst.gridgrind.core.result.AnalysisSummaryReport;
import dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode;
import dev.erst.gridgrind.core.vocabulary.AnalysisSeverity;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.TableName;
import dev.erst.gridgrind.engine.formulas.ComparisonFormulas;
import dev.erst.gridgrind.engine.tables.TableAdapter;
import java.io.IOException;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Tests for document-level validation analysis aggregation. */
class ExcelDocumentAnalyzerTest {
  @Test
  void dataValidationHealthHonorsSelectedSheets() throws IOException {
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      workbook.getOrCreateSheet("Budget");
      workbook.getOrCreateSheet("Archive");
      workbook
          .sheet("Budget")
          .metadata()
          .setDataValidation(
              "A1:A3",
              new ResolvedDataValidation(
                  new DataValidationRuleSpec.FormulaList(
                      ComparisonFormulas.normalizeFormula("#REF!", "formula")),
                  false,
                  false,
                  Optional.empty(),
                  Optional.empty()));
      workbook
          .sheet("Archive")
          .metadata()
          .setDataValidation(
              "A1:A3",
              new ResolvedDataValidation(
                  new DataValidationRuleSpec.ExplicitList(List.of("Queued", "Done")),
                  false,
                  false,
                  Optional.empty(),
                  Optional.empty()));

      DocumentAnalyzer analyzer = new DocumentAnalyzer();
      DataValidationHealthReport selected =
          analyzer.dataValidationHealth(
              workbook,
              new dev.erst.gridgrind.core.plan.SheetSelector.ByNames(
                  List.of("Budget").stream()
                      .map(dev.erst.gridgrind.core.vocabulary.SheetName::new)
                      .toList()));

      assertEquals(1, selected.checkedValidationCount());
      assertEquals(1, selected.summary().totalCount());
      assertEquals(1, selected.summary().errorCount());
      assertEquals(0, selected.summary().warningCount());
      assertEquals(0, selected.summary().infoCount());
      assertEquals(
          "Budget",
          ((AnalysisLocationReport.Range) selected.findings().getFirst().location())
              .sheetName()
              .value());
    }
  }

  @Test
  void autofilterAndTableHealthHonorTheirSelections() throws IOException {
    try (XssfWorkbook workbook = XssfWorkbooks.create()) {
      XssfSheet budget = workbook.getOrCreateSheet("Budget");
      budget.cells().setCell("A1", new CellValue.Text("Owner"));
      budget.cells().setCell("B1", new CellValue.Text("Task"));
      budget.cells().setCell("A2", new CellValue.Text("Ada"));
      budget.cells().setCell("B2", new CellValue.Text("Queue"));
      budget.cells().setCell("A3", new CellValue.Text("Lin"));
      budget.cells().setCell("B3", new CellValue.Text("Pack"));
      new TableAdapter()
          .setTable(
              workbook,
              new ResolvedTable(
                  new TableName("BudgetQueue"),
                  new SheetName("Budget"),
                  new RangeRef("A1:B3"),
                  false,
                  new TableStyleSpec.None()));
      budget.xssfSheet().getTables().getFirst().getCTTable().getAutoFilter().setRef("A1:B2");

      XssfSheet archive = workbook.getOrCreateSheet("Archive");
      archive.cells().setCell("D1", new CellValue.Text(""));
      archive.cells().setCell("E1", new CellValue.Text(""));
      archive.cells().setCell("D2", new CellValue.Text("Old"));
      archive.cells().setCell("E2", new CellValue.Text("Done"));
      archive.xssfSheet().getCTWorksheet().addNewAutoFilter().setRef("D1:E2");

      DocumentAnalyzer analyzer = new DocumentAnalyzer();
      AutofilterHealthReport autofilterHealth =
          analyzer.autofilterHealth(
              workbook,
              new dev.erst.gridgrind.core.plan.SheetSelector.ByNames(
                  List.of("Archive").stream()
                      .map(dev.erst.gridgrind.core.vocabulary.SheetName::new)
                      .toList()));
      TableHealthReport tableHealth =
          analyzer.tableHealth(
              workbook,
              new TableSelector.ByNames(
                  List.of("BudgetQueue").stream()
                      .map(dev.erst.gridgrind.core.vocabulary.TableName::new)
                      .toList()));

      assertEquals(1, autofilterHealth.checkedAutofilterCount());
      assertEquals(1, autofilterHealth.summary().warningCount());
      assertEquals(
          AnalysisFindingCode.AUTOFILTER_MISSING_HEADER_ROW,
          autofilterHealth.findings().getFirst().code());

      assertEquals(1, tableHealth.checkedTableCount());
      assertEquals(0, tableHealth.summary().totalCount());

      AutofilterHealthReport budgetAutofilterHealth =
          analyzer.autofilterHealth(
              workbook,
              new dev.erst.gridgrind.core.plan.SheetSelector.ByNames(
                  List.of("Budget").stream()
                      .map(dev.erst.gridgrind.core.vocabulary.SheetName::new)
                      .toList()));
      assertEquals(1, budgetAutofilterHealth.summary().warningCount());
      assertEquals(
          AnalysisFindingCode.AUTOFILTER_TABLE_MISMATCH,
          budgetAutofilterHealth.findings().getFirst().code());
    }
  }

  @Test
  void summaryCountsAllSeverities() {
    AnalysisSummaryReport summary =
        DocumentAnalyzer.summarizeFindings(
            List.of(
                finding(AnalysisSeverity.ERROR, "A1"),
                finding(AnalysisSeverity.WARNING, "A2"),
                finding(AnalysisSeverity.INFO, "A3")));

    assertEquals(3, summary.totalCount());
    assertEquals(1, summary.errorCount());
    assertEquals(1, summary.warningCount());
    assertEquals(1, summary.infoCount());
  }

  private static AnalysisFindingReport finding(AnalysisSeverity severity, String range) {
    return new AnalysisFindingReport(
        AnalysisFindingCode.DATA_VALIDATION_UNSUPPORTED_RULE,
        severity,
        "Finding",
        "Detail",
        new AnalysisLocationReport.Range(
            new dev.erst.gridgrind.core.vocabulary.SheetLabel("Budget"), range),
        List.of(range));
  }
}
