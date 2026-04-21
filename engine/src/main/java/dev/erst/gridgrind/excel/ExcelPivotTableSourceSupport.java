package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Set;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** Authoring-source lookup and source-column normalization helpers for pivot tables. */
final class ExcelPivotTableSourceSupport {
  private ExcelPivotTableSourceSupport() {}

  static ResolvedAuthoringSource resolveAuthoringSource(
      ExcelWorkbook workbook, ExcelPivotTableDefinition.Source source) {
    return switch (source) {
      case ExcelPivotTableDefinition.Source.Range range -> {
        XSSFSheet sheet = requiredSheet(workbook, range.sheetName());
        AreaReference area = ExcelPivotTableIdentitySupport.contiguousArea(range.range(), "range");
        yield ResolvedAuthoringSource.range(sheet, area);
      }
      case ExcelPivotTableDefinition.Source.NamedRange namedRange -> {
        List<Name> matches = matchingNamedRanges(workbook.xssfWorkbook(), namedRange.name(), null);
        if (matches.isEmpty()) {
          throw new IllegalArgumentException(
              "pivot source named range not found: " + namedRange.name());
        }
        if (matches.size() > 1) {
          throw new IllegalArgumentException(
              "pivot source named range is ambiguous: " + namedRange.name());
        }
        Name resolved = matches.getFirst();
        AreaReference area = namedRangeArea(resolved);
        XSSFSheet sheet = requiredSheet(workbook, sourceSheetName(area, resolved, null));
        yield ResolvedAuthoringSource.namedRange(sheet, area, resolved);
      }
      case ExcelPivotTableDefinition.Source.Table table -> {
        XSSFTable resolved = requiredTableByName(workbook, table.name());
        XSSFSheet sheet = requiredSheet(workbook, resolved.getSheetName());
        yield ResolvedAuthoringSource.table(
            sheet,
            new AreaReference(
                resolved.getStartCellReference(),
                resolved.getEndCellReference(),
                org.apache.poi.ss.SpreadsheetVersion.EXCEL2007),
            resolved);
      }
    };
  }

  static XSSFTable requiredTableByName(ExcelWorkbook workbook, String name) {
    XSSFTable table = tableByName(workbook.xssfWorkbook(), name, null);
    if (table == null) {
      throw new IllegalArgumentException("pivot source table not found: " + name);
    }
    return table;
  }

  static XSSFTable tableByName(XSSFWorkbook workbook, String name, String preferredSheetName) {
    String expected = name.toUpperCase(Locale.ROOT);
    XSSFTable match = null;
    for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
      XSSFSheet sheet = workbook.getSheetAt(sheetIndex);
      for (XSSFTable table : sheet.getTables()) {
        if (!table.getName().toUpperCase(Locale.ROOT).equals(expected)) {
          continue;
        }
        if (preferredSheetName != null && !sheet.getSheetName().equals(preferredSheetName)) {
          continue;
        }
        if (match != null) {
          throw new IllegalArgumentException("pivot source table is ambiguous: " + name);
        }
        match = table;
      }
    }
    return match;
  }

  static List<Name> matchingNamedRanges(XSSFWorkbook workbook, String name, String sheetNameHint) {
    List<Name> matches = new ArrayList<>();
    for (Name candidate : workbook.getAllNames()) {
      if (!candidate.getNameName().equalsIgnoreCase(name)) {
        continue;
      }
      if (sheetNameHint != null && candidate.getSheetIndex() >= 0) {
        String candidateSheetName = workbook.getSheetName(candidate.getSheetIndex());
        if (!sheetNameHint.equals(candidateSheetName)) {
          continue;
        }
      }
      matches.add(candidate);
    }
    return List.copyOf(matches);
  }

  static AreaReference namedRangeArea(Name namedRange) {
    String formula = java.util.Objects.requireNonNullElse(namedRange.getRefersToFormula(), "");
    if (formula.isBlank()) {
      throw new IllegalArgumentException(
          "pivot source named range " + namedRange.getNameName() + " does not define an area");
    }
    try {
      return new AreaReference(
          ExcelNamedRangeTargets.normalizeAreaFormulaForPoi(formula),
          org.apache.poi.ss.SpreadsheetVersion.EXCEL2007);
    } catch (RuntimeException exception) {
      throw new IllegalArgumentException(
          "pivot source named range "
              + namedRange.getNameName()
              + " does not define a rectangular area",
          exception);
    }
  }

  static String sourceSheetName(AreaReference area, Name namedRange, String fallbackSheetName) {
    String areaSheetName = area.getFirstCell().getSheetName();
    if (areaSheetName != null) {
      return areaSheetName;
    }
    if (namedRange != null && namedRange.getSheetIndex() >= 0) {
      return namedRange.getSheetName();
    }
    if (fallbackSheetName != null && !fallbackSheetName.isBlank()) {
      return fallbackSheetName;
    }
    throw new IllegalArgumentException("pivot source area does not identify its sheet");
  }

  static SourceColumns sourceColumns(XSSFSheet sheet, AreaReference area, String description) {
    CellReference firstCell = area.getFirstCell();
    CellReference lastCell = area.getLastCell();
    if (lastCell.getRow() <= firstCell.getRow()) {
      throw new IllegalArgumentException(
          "pivot source " + description + " must include a header row plus at least one data row");
    }
    Row headerRow = sheet.getRow(firstCell.getRow());
    if (headerRow == null) {
      throw new IllegalArgumentException(
          "pivot source " + description + " is missing its header row");
    }

    List<SourceColumn> columns = new ArrayList<>();
    Set<String> seenNames = new java.util.LinkedHashSet<>();
    for (int columnIndex = firstCell.getCol(); columnIndex <= lastCell.getCol(); columnIndex++) {
      var cell = headerRow.getCell(columnIndex);
      if (cell == null || cell.getCellType() != org.apache.poi.ss.usermodel.CellType.STRING) {
        throw new IllegalArgumentException(
            "pivot source " + description + " header cells must all be strings");
      }
      String name = cell.getStringCellValue();
      if (name.isBlank()) {
        throw new IllegalArgumentException(
            "pivot source " + description + " contains a blank header cell");
      }
      String key = name.toUpperCase(Locale.ROOT);
      if (!seenNames.add(key)) {
        throw new IllegalArgumentException(
            "pivot source "
                + description
                + " header row must be unique case-insensitively: "
                + name);
      }
      columns.add(new SourceColumn(name, columnIndex - firstCell.getCol()));
    }
    return new SourceColumns(columns);
  }

  static String sourceColumnName(List<String> sourceColumnNames, int sourceColumnIndex) {
    if (sourceColumnIndex < 0 || sourceColumnIndex >= sourceColumnNames.size()) {
      throw new IllegalArgumentException(
          "pivot source column index is out of bounds: " + sourceColumnIndex);
    }
    return sourceColumnNames.get(sourceColumnIndex);
  }

  static XSSFSheet requiredSheet(ExcelWorkbook workbook, String sheetName) {
    XSSFSheet sheet = workbook.xssfWorkbook().getSheet(sheetName);
    if (sheet == null) {
      throw new SheetNotFoundException(sheetName);
    }
    return sheet;
  }
}
