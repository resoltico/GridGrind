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
