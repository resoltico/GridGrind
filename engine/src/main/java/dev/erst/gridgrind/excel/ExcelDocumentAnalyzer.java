package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.pivot.PivotTableHealthReport;
import dev.erst.gridgrind.core.model.pivot.PivotTableReport;
import dev.erst.gridgrind.core.model.table.AutofilterHealthReport;
import dev.erst.gridgrind.core.model.table.TableEntryReport;
import dev.erst.gridgrind.core.model.table.TableHealthReport;
import dev.erst.gridgrind.core.model.validation.ConditionalFormattingHealthReport;
import dev.erst.gridgrind.core.model.validation.DataValidationHealthReport;
import dev.erst.gridgrind.core.plan.PivotTableSelector;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.core.result.AnalysisFindingReport;
import dev.erst.gridgrind.core.result.AnalysisSummaryReport;
import dev.erst.gridgrind.core.vocabulary.AnalysisSeverity;
import dev.erst.gridgrind.excel.pivot.ExcelPivotTableController;
import dev.erst.gridgrind.excel.validation.ExcelDataValidationController;
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
  DataValidationHealthReport dataValidationHealth(ExcelWorkbook workbook, SheetSelector selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");

    int checkedValidationCount = 0;
    List<AnalysisFindingReport> findings = new ArrayList<>();
    for (String sheetName : selectSheets(workbook, selection)) {
      ExcelSheet sheet = workbook.sheet(sheetName);
      checkedValidationCount += dataValidationController.dataValidationCount(sheet.xssfSheet());
      findings.addAll(
          dataValidationController.dataValidationHealthFindings(sheetName, sheet.xssfSheet()));
    }
    AnalysisSummaryReport summary = summarizeFindings(findings);
    return new DataValidationHealthReport(checkedValidationCount, summary, List.copyOf(findings));
  }

  /** Returns conditional-formatting-health findings for the selected sheets. */
  ConditionalFormattingHealthReport conditionalFormattingHealth(
      ExcelWorkbook workbook, SheetSelector selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");

    int checkedConditionalFormattingBlockCount = 0;
    List<AnalysisFindingReport> findings = new ArrayList<>();
    for (String sheetName : selectSheets(workbook, selection)) {
      ExcelSheet sheet = workbook.sheet(sheetName);
      checkedConditionalFormattingBlockCount +=
          conditionalFormattingController.conditionalFormattingBlockCount(sheet.xssfSheet());
      findings.addAll(
          conditionalFormattingController.conditionalFormattingHealthFindings(
              sheetName, sheet.xssfSheet()));
    }
    return new ConditionalFormattingHealthReport(
        checkedConditionalFormattingBlockCount, summarizeFindings(findings), List.copyOf(findings));
  }

  /** Returns autofilter-health findings for the selected sheets. */
  AutofilterHealthReport autofilterHealth(ExcelWorkbook workbook, SheetSelector selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");

    List<TableEntryReport> allTables = tableController.tables(workbook, new TableSelector.All());
    int checkedAutofilterCount = 0;
    List<AnalysisFindingReport> findings = new ArrayList<>();
    for (String sheetName : selectSheets(workbook, selection)) {
      ExcelSheet sheet = workbook.sheet(sheetName);
      checkedAutofilterCount += autofilterController.sheetAutofilterCount(sheet.xssfSheet());
      checkedAutofilterCount += tableController.tableAutofilterCount(workbook, sheetName);
      List<TableEntryReport> tablesOnSheet =
          allTables.stream().filter(table -> table.sheetName().value().equals(sheetName)).toList();
      findings.addAll(
          autofilterController.sheetAutofilterHealthFindings(
              sheetName, sheet.xssfSheet(), tablesOnSheet));
      findings.addAll(tableController.tableAutofilterHealthFindings(workbook, sheetName));
    }
    return new AutofilterHealthReport(
        checkedAutofilterCount, summarizeFindings(findings), List.copyOf(findings));
  }

  /** Returns table-health findings for the selected workbook-global tables. */
  TableHealthReport tableHealth(ExcelWorkbook workbook, TableSelector selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");

    List<TableEntryReport> selectedTables = tableController.tables(workbook, selection);
    List<AnalysisFindingReport> findings = tableController.tableHealthFindings(workbook, selection);
    return new TableHealthReport(
        selectedTables.size(), summarizeFindings(findings), List.copyOf(findings));
  }

  /** Returns pivot-table-health findings for the selected workbook-global pivot tables. */
  PivotTableHealthReport pivotTableHealth(ExcelWorkbook workbook, PivotTableSelector selection) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(selection, "selection must not be null");

    List<PivotTableReport> selectedPivots = pivotTableController.pivotTables(workbook, selection);
    List<AnalysisFindingReport> findings =
        pivotTableController.pivotTableHealthFindings(workbook, selection);
    return new PivotTableHealthReport(
        selectedPivots.size(), summarizeFindings(findings), List.copyOf(findings));
  }

  private List<String> selectSheets(ExcelWorkbook workbook, SheetSelector selection) {
    return switch (selection) {
      case SheetSelector.All _ -> workbook.sheets().sheetNames();
      case SheetSelector.ByNames selected ->
          selected.names().stream()
              .map(dev.erst.gridgrind.core.vocabulary.SheetName::value)
              .toList();
      case SheetSelector.ByName selected -> List.of(selected.name().value());
    };
  }

  /** Returns analysis summary counts derived from the provided findings. */
  static AnalysisSummaryReport summarizeFindings(List<AnalysisFindingReport> findings) {
    int errorCount = 0;
    int warningCount = 0;
    int infoCount = 0;
    for (AnalysisFindingReport finding : findings) {
      if (finding.severity() == AnalysisSeverity.ERROR) {
        errorCount++;
        continue;
      }
      if (finding.severity() == AnalysisSeverity.WARNING) {
        warningCount++;
        continue;
      }
      infoCount++;
    }
    return new AnalysisSummaryReport(findings.size(), errorCount, warningCount, infoCount);
  }
}
