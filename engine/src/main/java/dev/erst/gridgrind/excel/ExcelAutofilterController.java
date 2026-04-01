package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/** Reads, writes, and analyzes sheet-owned autofilter structures on one XSSF sheet. */
final class ExcelAutofilterController {
  /** Creates or replaces one sheet-level autofilter range. */
  void setSheetAutofilter(XSSFSheet sheet, String range) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(range, "range must not be null");

    ExcelRange targetRange = ExcelRange.parse(range);
    if (ExcelSheetStructureSupport.headerRowMissing(sheet, targetRange)) {
      throw new IllegalArgumentException(
          "autofilter range must include a nonblank header row: "
              + ExcelSheetStructureSupport.formatRange(targetRange));
    }
    requireNoTableOverlap(sheet, targetRange);
    sheet.setAutoFilter(ExcelSheetStructureSupport.toCellRangeAddress(targetRange));
  }

  /** Clears the sheet-level autofilter range on one sheet. */
  void clearSheetAutofilter(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    if (sheet.getCTWorksheet().isSetAutoFilter()) {
      sheet.getCTWorksheet().unsetAutoFilter();
    }
    List<Name> filterDatabaseNames = new ArrayList<>();
    int sheetIndex = sheet.getWorkbook().getSheetIndex(sheet);
    for (Name name : sheet.getWorkbook().getAllNames()) {
      if (name.getSheetIndex() == sheetIndex
          && "_XLNM._FILTERDATABASE".equalsIgnoreCase(name.getNameName())) {
        filterDatabaseNames.add(name);
      }
    }
    for (Name filterDatabaseName : filterDatabaseNames) {
      sheet.getWorkbook().removeName(filterDatabaseName);
    }
  }

  /** Returns the sheet-owned autofilter metadata present on one sheet. */
  List<ExcelAutofilterSnapshot> sheetOwnedAutofilters(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    if (!sheet.getCTWorksheet().isSetAutoFilter()) {
      return List.of();
    }
    String rawRange =
        Objects.requireNonNullElse(sheet.getCTWorksheet().getAutoFilter().getRef(), "");
    return List.of(new ExcelAutofilterSnapshot.SheetOwned(rawRange));
  }

  /** Returns the number of sheet-owned autofilters currently present on one sheet. */
  int sheetAutofilterCount(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    return sheet.getCTWorksheet().isSetAutoFilter() ? 1 : 0;
  }

  /** Returns derived health findings for the sheet-owned autofilter on one sheet. */
  List<WorkbookAnalysis.AnalysisFinding> sheetAutofilterHealthFindings(
      String sheetName, XSSFSheet sheet, List<ExcelTableSnapshot> tablesOnSheet) {
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(tablesOnSheet, "tablesOnSheet must not be null");

    if (!sheet.getCTWorksheet().isSetAutoFilter()) {
      return List.of();
    }

    String rawRange =
        Objects.requireNonNullElse(sheet.getCTWorksheet().getAutoFilter().getRef(), "");
    ExcelRange targetRange = ExcelSheetStructureSupport.parseRangeOrNull(rawRange);
    List<WorkbookAnalysis.AnalysisFinding> findings = new ArrayList<>();
    if (targetRange == null) {
      findings.add(
          new WorkbookAnalysis.AnalysisFinding(
              WorkbookAnalysis.AnalysisFindingCode.AUTOFILTER_INVALID_RANGE,
              WorkbookAnalysis.AnalysisSeverity.ERROR,
              "Autofilter range is invalid",
              "Sheet-owned autofilter range could not be parsed.",
              new WorkbookAnalysis.AnalysisLocation.Sheet(sheetName),
              List.of(rawRange)));
      return List.copyOf(findings);
    }

    String normalizedRange = ExcelSheetStructureSupport.formatRange(targetRange);
    WorkbookAnalysis.AnalysisLocation.Range location =
        new WorkbookAnalysis.AnalysisLocation.Range(sheetName, normalizedRange);
    if (ExcelSheetStructureSupport.headerRowMissing(sheet, targetRange)) {
      findings.add(
          new WorkbookAnalysis.AnalysisFinding(
              WorkbookAnalysis.AnalysisFindingCode.AUTOFILTER_MISSING_HEADER_ROW,
              WorkbookAnalysis.AnalysisSeverity.WARNING,
              "Autofilter header row is blank",
              "Sheet-owned autofilter range does not contain a nonblank header row.",
              location,
              List.of(normalizedRange)));
    }

    List<String> overlappingTables = new ArrayList<>();
    for (ExcelTableSnapshot table : tablesOnSheet) {
      ExcelRange tableRange = ExcelSheetStructureSupport.parseRangeOrNull(table.range());
      if (tableRange != null && ExcelSheetStructureSupport.intersects(targetRange, tableRange)) {
        overlappingTables.add(table.name() + "@" + table.range());
      }
    }
    if (!overlappingTables.isEmpty()) {
      List<String> evidence = new ArrayList<>();
      evidence.add(normalizedRange);
      evidence.addAll(overlappingTables);
      findings.add(
          new WorkbookAnalysis.AnalysisFinding(
              WorkbookAnalysis.AnalysisFindingCode.AUTOFILTER_TABLE_MISMATCH,
              WorkbookAnalysis.AnalysisSeverity.WARNING,
              "Sheet autofilter overlaps a table range",
              "Sheet-owned autofilter metadata overlaps one or more table ranges."
                  + " Table-owned filters should be managed by table definitions instead.",
              location,
              List.copyOf(evidence)));
    }
    return List.copyOf(findings);
  }

  private void requireNoTableOverlap(XSSFSheet sheet, ExcelRange targetRange) {
    List<String> overlaps = new ArrayList<>();
    for (var table : sheet.getTables()) {
      ExcelRange tableRange =
          ExcelSheetStructureSupport.parseRangeOrNull(
              Objects.requireNonNullElse(table.getCTTable().getRef(), ""));
      if (tableRange != null && ExcelSheetStructureSupport.intersects(targetRange, tableRange)) {
        overlaps.add(table.getName() + "@" + table.getCTTable().getRef());
      }
    }
    if (!overlaps.isEmpty()) {
      throw new IllegalArgumentException(
          "sheet-level autofilter range must not overlap an existing table range: "
              + String.join(", ", overlaps));
    }
  }
}
