package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Derives document-intelligence findings from reusable workbook and sheet facts. */
final class ExcelDocumentAnalyzer {
  private final ExcelDataValidationController dataValidationController;
  private final ExcelConditionalFormattingController conditionalFormattingController;
  private final ExcelAutofilterController autofilterController;
  private final ExcelTableController tableController;
  private final ExcelPivotTableController pivotTableController;

  ExcelDocumentAnalyzer() {
    this(
        new ExcelDataValidationController(),
        new ExcelConditionalFormattingController(),
        new ExcelAutofilterController(),
        new ExcelTableController(),
        new ExcelPivotTableController());
  }

  ExcelDocumentAnalyzer(
      ExcelDataValidationController dataValidationController,
      ExcelConditionalFormattingController conditionalFormattingController,
      ExcelAutofilterController autofilterController,
      ExcelTableController tableController,
      ExcelPivotTableController pivotTableController) {
    this.dataValidationController =
        Objects.requireNonNull(
            dataValidationController, "dataValidationController must not be null");
    this.conditionalFormattingController =
        Objects.requireNonNull(
            conditionalFormattingController, "conditionalFormattingController must not be null");
    this.autofilterController =
        Objects.requireNonNull(autofilterController, "autofilterController must not be null");
    this.tableController =
        Objects.requireNonNull(tableController, "tableController must not be null");
    this.pivotTableController =
        Objects.requireNonNull(pivotTableController, "pivotTableController must not be null");
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

  /** Returns conditional-formatting-health findings for the selected sheets. */
  WorkbookAnalysis.ConditionalFormattingHealth conditionalFormattingHealth(
      ExcelWorkbook workbook, ExcelSheetSelection selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");

    int checkedConditionalFormattingBlockCount = 0;
    List<WorkbookAnalysis.AnalysisFinding> findings = new ArrayList<>();
    for (String sheetName : selectSheets(workbook, selection)) {
      ExcelSheet sheet = workbook.sheet(sheetName);
      checkedConditionalFormattingBlockCount +=
          conditionalFormattingController.conditionalFormattingBlockCount(sheet.xssfSheet());
      findings.addAll(
          conditionalFormattingController.conditionalFormattingHealthFindings(
              sheetName, sheet.xssfSheet()));
    }
    return new WorkbookAnalysis.ConditionalFormattingHealth(
        checkedConditionalFormattingBlockCount, summarizeFindings(findings), List.copyOf(findings));
  }

  /** Returns autofilter-health findings for the selected sheets. */
  WorkbookAnalysis.AutofilterHealth autofilterHealth(
      ExcelWorkbook workbook, ExcelSheetSelection selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");

    List<ExcelTableSnapshot> allTables =
        tableController.tables(workbook, new ExcelTableSelection.All());
    int checkedAutofilterCount = 0;
    List<WorkbookAnalysis.AnalysisFinding> findings = new ArrayList<>();
    for (String sheetName : selectSheets(workbook, selection)) {
      ExcelSheet sheet = workbook.sheet(sheetName);
      checkedAutofilterCount += autofilterController.sheetAutofilterCount(sheet.xssfSheet());
      checkedAutofilterCount += tableController.tableAutofilterCount(workbook, sheetName);
      List<ExcelTableSnapshot> tablesOnSheet =
          allTables.stream().filter(table -> table.sheetName().equals(sheetName)).toList();
      findings.addAll(
          autofilterController.sheetAutofilterHealthFindings(
              sheetName, sheet.xssfSheet(), tablesOnSheet));
      findings.addAll(tableController.tableAutofilterHealthFindings(workbook, sheetName));
    }
    return new WorkbookAnalysis.AutofilterHealth(
        checkedAutofilterCount, summarizeFindings(findings), List.copyOf(findings));
  }

  /** Returns table-health findings for the selected workbook-global tables. */
  WorkbookAnalysis.TableHealth tableHealth(ExcelWorkbook workbook, ExcelTableSelection selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");

    List<ExcelTableSnapshot> selectedTables = tableController.tables(workbook, selection);
    List<WorkbookAnalysis.AnalysisFinding> findings =
        tableController.tableHealthFindings(workbook, selection);
    return new WorkbookAnalysis.TableHealth(
        selectedTables.size(), summarizeFindings(findings), List.copyOf(findings));
  }

  /** Returns pivot-table-health findings for the selected workbook-global pivot tables. */
  WorkbookAnalysis.PivotTableHealth pivotTableHealth(
      ExcelWorkbook workbook, ExcelPivotTableSelection selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");

    List<ExcelPivotTableSnapshot> selectedPivots =
        pivotTableController.pivotTables(workbook, selection);
    List<WorkbookAnalysis.AnalysisFinding> findings =
        pivotTableController.pivotTableHealthFindings(workbook, selection);
    return new WorkbookAnalysis.PivotTableHealth(
        selectedPivots.size(), summarizeFindings(findings), List.copyOf(findings));
  }

  private List<String> selectSheets(ExcelWorkbook workbook, ExcelSheetSelection selection) {
    return switch (selection) {
      case ExcelSheetSelection.All _ -> workbook.sheetNames();
      case ExcelSheetSelection.Selected selected -> List.copyOf(selected.sheetNames());
    };
  }

  /** Returns analysis summary counts derived from the provided findings. */
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
