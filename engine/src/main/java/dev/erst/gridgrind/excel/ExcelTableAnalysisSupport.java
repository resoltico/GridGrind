package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.table.TableEntryReport;
import dev.erst.gridgrind.core.model.table.TableStyleReport;
import dev.erst.gridgrind.core.vocabulary.AnalysisFindingCode;
import dev.erst.gridgrind.core.vocabulary.AnalysisSeverity;
import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;

/** Shared derived-analysis helpers for factual workbook table metadata. */
final class ExcelTableAnalysisSupport {
  private ExcelTableAnalysisSupport() {}

  /** Returns derived health findings for one factual table snapshot. */
  static List<WorkbookAnalysis.AnalysisFinding> tableFindings(
      ExcelWorkbook workbook, TableEntryReport table) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(table, "table must not be null");

    List<WorkbookAnalysis.AnalysisFinding> findings = new ArrayList<>();
    Optional<ExcelRange> parsedRange =
        ExcelSheetStructureSupport.parseOptionalRange(table.range().value());
    WorkbookAnalysis.AnalysisLocation location = tableLocation(table, parsedRange);
    if (parsedRange.isEmpty()) {
      return brokenReferenceFindings(table, location);
    }

    ExcelRange range = parsedRange.orElseThrow();
    addRowCountFinding(findings, table, range, location);
    addHeaderFindings(findings, table, location);
    addStyleFinding(findings, workbook, table, location);
    return List.copyOf(findings);
  }

  private static WorkbookAnalysis.AnalysisLocation tableLocation(
      TableEntryReport table, Optional<ExcelRange> parsedRange) {
    return parsedRange.isEmpty()
        ? new WorkbookAnalysis.AnalysisLocation.Sheet(table.sheetName().value())
        : new WorkbookAnalysis.AnalysisLocation.Range(
            table.sheetName().value(),
            ExcelSheetStructureSupport.formatRange(parsedRange.orElseThrow()));
  }

  private static List<WorkbookAnalysis.AnalysisFinding> brokenReferenceFindings(
      TableEntryReport table, WorkbookAnalysis.AnalysisLocation location) {
    return List.of(
        new WorkbookAnalysis.AnalysisFinding(
            AnalysisFindingCode.TABLE_BROKEN_REFERENCE,
            AnalysisSeverity.ERROR,
            "Table range is invalid",
            "Table range could not be parsed from workbook metadata.",
            location,
            List.of(table.name().value(), table.range().value())));
  }

  private static void addRowCountFinding(
      List<WorkbookAnalysis.AnalysisFinding> findings,
      TableEntryReport table,
      ExcelRange range,
      WorkbookAnalysis.AnalysisLocation location) {
    int minimumRows = table.structure().headerRowCount() + table.structure().totalsRowCount() + 1;
    if (table.structure().headerRowCount() < 1 || range.rowCount() < minimumRows) {
      findings.add(
          new WorkbookAnalysis.AnalysisFinding(
              AnalysisFindingCode.TABLE_BROKEN_REFERENCE,
              AnalysisSeverity.ERROR,
              "Table range does not match table row counts",
              "Table metadata requires at least "
                  + minimumRows
                  + " rows but the stored range contains only "
                  + range.rowCount()
                  + ".",
              location,
              List.of(table.range().value())));
    }
  }

  private static void addHeaderFindings(
      List<WorkbookAnalysis.AnalysisFinding> findings,
      TableEntryReport table,
      WorkbookAnalysis.AnalysisLocation location) {
    List<String> duplicateHeaders = new ArrayList<>();
    List<String> blankHeaders = new ArrayList<>();
    Set<String> seen = new LinkedHashSet<>();
    for (String header : table.structure().columnNames()) {
      if (header.isBlank()) {
        blankHeaders.add(header);
        continue;
      }
      if (!seen.add(header.toUpperCase(Locale.ROOT))) {
        duplicateHeaders.add(header);
      }
    }
    if (!blankHeaders.isEmpty()) {
      findings.add(
          new WorkbookAnalysis.AnalysisFinding(
              AnalysisFindingCode.TABLE_BLANK_HEADER,
              AnalysisSeverity.ERROR,
              "Table contains blank header cells",
              "Table column headers must be nonblank.",
              location,
              List.of(table.name().value(), table.range().value())));
    }
    if (!duplicateHeaders.isEmpty()) {
      findings.add(
          new WorkbookAnalysis.AnalysisFinding(
              AnalysisFindingCode.TABLE_DUPLICATE_HEADER,
              AnalysisSeverity.ERROR,
              "Table contains duplicate header cells",
              "Table column headers must be unique (case-insensitive).",
              location,
              List.copyOf(duplicateHeaders)));
    }
  }

  private static void addStyleFinding(
      List<WorkbookAnalysis.AnalysisFinding> findings,
      ExcelWorkbook workbook,
      TableEntryReport table,
      WorkbookAnalysis.AnalysisLocation location) {
    switch (table.style()) {
      case TableStyleReport.None _ -> {}
      case TableStyleReport.Named named -> {
        if (named.name().isBlank()
            || workbook.xssfWorkbook().getStylesSource().getTableStyle(named.name()) == null) {
          findings.add(
              new WorkbookAnalysis.AnalysisFinding(
                  AnalysisFindingCode.TABLE_STYLE_MISMATCH,
                  AnalysisSeverity.WARNING,
                  "Table style does not resolve",
                  "Table style metadata refers to a style name that the workbook does not define.",
                  location,
                  List.of(named.name())));
        }
      }
    }
  }

  /** Returns derived overlap findings for one factual table snapshot against a peer set. */
  static List<WorkbookAnalysis.AnalysisFinding> overlapFindings(
      TableEntryReport table, List<TableEntryReport> allTables) {
    Objects.requireNonNull(table, "table must not be null");
    Objects.requireNonNull(allTables, "allTables must not be null");

    Optional<ExcelRange> tableRange =
        ExcelSheetStructureSupport.parseOptionalRange(table.range().value());
    if (tableRange.isEmpty()) {
      return List.of();
    }
    List<String> overlaps = new ArrayList<>();
    for (TableEntryReport other : allTables) {
      if (other.equals(table) || !other.sheetName().value().equals(table.sheetName().value())) {
        continue;
      }
      Optional<ExcelRange> otherRange =
          ExcelSheetStructureSupport.parseOptionalRange(other.range().value());
      if (otherRange.isPresent()
          && ExcelSheetStructureSupport.intersects(
              tableRange.orElseThrow(), otherRange.orElseThrow())) {
        overlaps.add(other.name().value() + "@" + other.range().value());
      }
    }
    if (overlaps.isEmpty()) {
      return List.of();
    }
    List<String> evidence = new ArrayList<>();
    evidence.add(table.name().value() + "@" + table.range().value());
    evidence.addAll(overlaps);
    return List.of(
        new WorkbookAnalysis.AnalysisFinding(
            AnalysisFindingCode.TABLE_OVERLAPPING_RANGE,
            AnalysisSeverity.ERROR,
            "Table range overlaps another table",
            "Table metadata overlaps one or more other tables on the same sheet.",
            new WorkbookAnalysis.AnalysisLocation.Range(
                table.sheetName().value(), table.range().value()),
            List.copyOf(evidence)));
  }

  /** Returns derived autofilter findings for one factual table snapshot on one sheet. */
  static List<WorkbookAnalysis.AnalysisFinding> tableAutofilterFindings(
      XSSFSheet sheet, TableEntryReport table) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(table, "table must not be null");

    Optional<ExcelRange> tableRange =
        ExcelSheetStructureSupport.parseOptionalRange(table.range().value());
    WorkbookAnalysis.AnalysisLocation location =
        tableRange.isEmpty()
            ? new WorkbookAnalysis.AnalysisLocation.Sheet(table.sheetName().value())
            : new WorkbookAnalysis.AnalysisLocation.Range(
                table.sheetName().value(),
                ExcelTableStructureSupport.expectedAutofilterRangeText(table));

    XSSFTable xssfTable = ExcelTableCatalogSupport.requiredTableByName(sheet, table.name().value());
    String rawFilterRange =
        Objects.requireNonNullElse(xssfTable.getCTTable().getAutoFilter().getRef(), "");
    Optional<ExcelRange> parsedFilterRange =
        ExcelSheetStructureSupport.parseOptionalRange(rawFilterRange);
    if (parsedFilterRange.isEmpty()) {
      return List.of(
          new WorkbookAnalysis.AnalysisFinding(
              AnalysisFindingCode.AUTOFILTER_INVALID_RANGE,
              AnalysisSeverity.ERROR,
              "Table autofilter range is invalid",
              "Table-owned autofilter range could not be parsed from workbook metadata.",
              new WorkbookAnalysis.AnalysisLocation.Sheet(table.sheetName().value()),
              List.of(table.name().value(), rawFilterRange)));
    }

    List<WorkbookAnalysis.AnalysisFinding> findings = new ArrayList<>();
    if (ExcelSheetStructureSupport.headerRowMissing(sheet, parsedFilterRange.orElseThrow())) {
      findings.add(
          new WorkbookAnalysis.AnalysisFinding(
              AnalysisFindingCode.AUTOFILTER_MISSING_HEADER_ROW,
              AnalysisSeverity.WARNING,
              "Table autofilter header row is blank",
              "Table-owned autofilter range does not contain a nonblank header row.",
              location,
              List.of(rawFilterRange, table.name().value())));
    }
    String expectedRange = ExcelTableStructureSupport.expectedAutofilterRangeText(table);
    if (!ExcelSheetStructureSupport.formatRange(parsedFilterRange.orElseThrow())
        .equals(expectedRange)) {
      findings.add(
          new WorkbookAnalysis.AnalysisFinding(
              AnalysisFindingCode.AUTOFILTER_TABLE_MISMATCH,
              AnalysisSeverity.WARNING,
              "Table autofilter does not match table range",
              "Table-owned autofilter range must match the table range excluding any totals row.",
              location,
              List.of(table.name().value(), rawFilterRange, expectedRange)));
    }
    return List.copyOf(findings);
  }
}
