package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Derives document-intelligence findings from reusable workbook and sheet facts. */
final class ExcelDocumentAnalyzer {
  private final ExcelDataValidationController dataValidationController;

  ExcelDocumentAnalyzer() {
    this(new ExcelDataValidationController());
  }

  ExcelDocumentAnalyzer(ExcelDataValidationController dataValidationController) {
    this.dataValidationController =
        Objects.requireNonNull(
            dataValidationController, "dataValidationController must not be null");
  }

  /** Returns data-validation-health findings for the selected sheets. */
  WorkbookAnalysis.DataValidationHealth dataValidationHealth(
      ExcelWorkbook workbook, ExcelSheetSelection selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");

    int checkedValidationCount = 0;
    List<WorkbookAnalysis.AnalysisFinding> findings = new ArrayList<>();
    for (String sheetName : selectSheets(workbook, selection)) {
      ExcelSheet sheet = workbook.sheet(sheetName);
      checkedValidationCount += dataValidationController.dataValidationCount(sheet.xssfSheet());
      findings.addAll(
          dataValidationController.dataValidationHealthFindings(sheetName, sheet.xssfSheet()));
    }
    WorkbookAnalysis.AnalysisSummary summary = summarizeFindings(findings);
    return new WorkbookAnalysis.DataValidationHealth(
        checkedValidationCount, summary, List.copyOf(findings));
  }

  private List<String> selectSheets(ExcelWorkbook workbook, ExcelSheetSelection selection) {
    return switch (selection) {
      case ExcelSheetSelection.All _ -> workbook.sheetNames();
      case ExcelSheetSelection.Selected selected -> List.copyOf(selected.sheetNames());
    };
  }

  static WorkbookAnalysis.AnalysisSummary summarizeFindings(
      List<WorkbookAnalysis.AnalysisFinding> findings) {
    int errorCount = 0;
    int warningCount = 0;
    int infoCount = 0;
    for (WorkbookAnalysis.AnalysisFinding finding : findings) {
      if (finding.severity() == WorkbookAnalysis.AnalysisSeverity.ERROR) {
        errorCount++;
        continue;
      }
      if (finding.severity() == WorkbookAnalysis.AnalysisSeverity.WARNING) {
        warningCount++;
        continue;
      }
      infoCount++;
    }
    return new WorkbookAnalysis.AnalysisSummary(
        findings.size(), errorCount, warningCount, infoCount);
  }
}
