package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.io.IOException;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for document-level validation analysis aggregation. */
class ExcelDocumentAnalyzerTest {
  @Test
  void dataValidationHealthHonorsSelectedSheets() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Budget");
      workbook.getOrCreateSheet("Archive");
      workbook
          .sheet("Budget")
          .setDataValidation(
              "A1:A3",
              new ExcelDataValidationDefinition(
                  new ExcelDataValidationRule.FormulaList("#REF!"), false, false, null, null));
      workbook
          .sheet("Archive")
          .setDataValidation(
              "A1:A3",
              new ExcelDataValidationDefinition(
                  new ExcelDataValidationRule.ExplicitList(List.of("Queued", "Done")),
                  false,
                  false,
                  null,
                  null));

      ExcelDocumentAnalyzer analyzer = new ExcelDocumentAnalyzer();
      WorkbookAnalysis.DataValidationHealth selected =
          analyzer.dataValidationHealth(
              workbook, new ExcelSheetSelection.Selected(List.of("Budget")));

      assertEquals(1, selected.checkedValidationCount());
      assertEquals(1, selected.summary().totalCount());
      assertEquals(1, selected.summary().errorCount());
      assertEquals(0, selected.summary().warningCount());
      assertEquals(0, selected.summary().infoCount());
      assertEquals(
          "Budget",
          ((WorkbookAnalysis.AnalysisLocation.Range) selected.findings().getFirst().location())
              .sheetName());
    }
  }

  @Test
  void autofilterAndTableHealthHonorTheirSelections() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      ExcelSheet budget = workbook.getOrCreateSheet("Budget");
      budget.setCell("A1", ExcelCellValue.text("Owner"));
      budget.setCell("B1", ExcelCellValue.text("Task"));
      budget.setCell("A2", ExcelCellValue.text("Ada"));
      budget.setCell("B2", ExcelCellValue.text("Queue"));
      budget.setCell("A3", ExcelCellValue.text("Lin"));
      budget.setCell("B3", ExcelCellValue.text("Pack"));
      workbook.setTable(
          new ExcelTableDefinition(
              "BudgetQueue", "Budget", "A1:B3", false, new ExcelTableStyle.None()));
      budget.xssfSheet().getTables().getFirst().getCTTable().getAutoFilter().setRef("A1:B2");

      ExcelSheet archive = workbook.getOrCreateSheet("Archive");
      archive.setCell("D1", ExcelCellValue.text(""));
      archive.setCell("E1", ExcelCellValue.text(""));
      archive.setCell("D2", ExcelCellValue.text("Old"));
      archive.setCell("E2", ExcelCellValue.text("Done"));
      archive.xssfSheet().getCTWorksheet().addNewAutoFilter().setRef("D1:E2");

      ExcelDocumentAnalyzer analyzer = new ExcelDocumentAnalyzer();
      WorkbookAnalysis.AutofilterHealth autofilterHealth =
          analyzer.autofilterHealth(workbook, new ExcelSheetSelection.Selected(List.of("Archive")));
      WorkbookAnalysis.TableHealth tableHealth =
          analyzer.tableHealth(workbook, new ExcelTableSelection.ByNames(List.of("BudgetQueue")));

      assertEquals(1, autofilterHealth.checkedAutofilterCount());
      assertEquals(1, autofilterHealth.summary().warningCount());
      assertEquals(
          WorkbookAnalysis.AnalysisFindingCode.AUTOFILTER_MISSING_HEADER_ROW,
          autofilterHealth.findings().getFirst().code());

      assertEquals(1, tableHealth.checkedTableCount());
      assertEquals(0, tableHealth.summary().totalCount());

      WorkbookAnalysis.AutofilterHealth budgetAutofilterHealth =
          analyzer.autofilterHealth(workbook, new ExcelSheetSelection.Selected(List.of("Budget")));
      assertEquals(1, budgetAutofilterHealth.summary().warningCount());
      assertEquals(
          WorkbookAnalysis.AnalysisFindingCode.AUTOFILTER_TABLE_MISMATCH,
          budgetAutofilterHealth.findings().getFirst().code());
    }
  }

  @Test
  void summaryCountsAllSeverities() {
    WorkbookAnalysis.AnalysisSummary summary =
        ExcelDocumentAnalyzer.summarizeFindings(
            List.of(
                finding(WorkbookAnalysis.AnalysisSeverity.ERROR, "A1"),
                finding(WorkbookAnalysis.AnalysisSeverity.WARNING, "A2"),
                finding(WorkbookAnalysis.AnalysisSeverity.INFO, "A3")));

    assertEquals(3, summary.totalCount());
    assertEquals(1, summary.errorCount());
    assertEquals(1, summary.warningCount());
    assertEquals(1, summary.infoCount());
  }

  private static WorkbookAnalysis.AnalysisFinding finding(
      WorkbookAnalysis.AnalysisSeverity severity, String range) {
    return new WorkbookAnalysis.AnalysisFinding(
        WorkbookAnalysis.AnalysisFindingCode.DATA_VALIDATION_UNSUPPORTED_RULE,
        severity,
        "Finding",
        "Detail",
        new WorkbookAnalysis.AnalysisLocation.Range("Budget", range),
        List.of(range));
  }
}
