package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelColumnSpan;
import dev.erst.gridgrind.excel.foundation.ExcelRowSpan;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidation;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidations;

/** Owns structural safety guards for row and column edits. */
final class ExcelRowColumnStructureGuardSupport {
  private ExcelRowColumnStructureGuardSupport() {}

  static void rejectAffectedRowStructuresForInsert(XSSFSheet sheet, int rowIndex) {
    for (XSSFTable table : sheet.getTables()) {
      ExcelRange range = parsedRange(table.getCTTable().getRef(), "table", table.getName());
      if (range.lastRow() >= rowIndex) {
        throw unsupportedStructure(
            "INSERT_ROWS",
            sheet,
            "table '" + table.getName() + "'",
            "row structural edits that would move tables are not supported");
      }
    }
    ExcelRange autofilter = sheetAutofilterRange(sheet);
    if (autofilter != null && autofilter.lastRow() >= rowIndex) {
      throw unsupportedStructure(
          "INSERT_ROWS",
          sheet,
          "sheet autofilter " + ExcelSheetStructureSupport.formatRange(autofilter),
          "row structural edits that would move sheet autofilters are not supported");
    }
  }

  static void rejectAffectedRowStructuresForDelete(XSSFSheet sheet, ExcelRowSpan rows) {
    for (XSSFTable table : sheet.getTables()) {
      ExcelRange range = parsedRange(table.getCTTable().getRef(), "table", table.getName());
      if (range.lastRow() >= rows.firstRowIndex()) {
        throw unsupportedStructure(
            "DELETE_ROWS",
            sheet,
            "table '" + table.getName() + "'",
            "row structural edits that would move or truncate tables are not supported");
      }
    }
    ExcelRange autofilter = sheetAutofilterRange(sheet);
    if (autofilter != null && autofilter.lastRow() >= rows.firstRowIndex()) {
      throw unsupportedStructure(
          "DELETE_ROWS",
          sheet,
          "sheet autofilter " + ExcelSheetStructureSupport.formatRange(autofilter),
          "row structural edits that would move or truncate sheet autofilters are not supported");
    }
    for (String validationRange : dataValidationRanges(sheet)) {
      ExcelRange range = parsedRange(validationRange, "data validation", validationRange);
      if (range.lastRow() >= rows.firstRowIndex()) {
        throw unsupportedStructure(
            "DELETE_ROWS",
            sheet,
            "data validation " + validationRange,
            "row structural edits that would move or truncate data validations are not supported");
      }
    }
    rejectDestructiveNamedRangesForRowDelete(sheet.getWorkbook(), sheet, rows);
  }

  static void rejectAffectedRowStructuresForShift(XSSFSheet sheet, ExcelRowSpan rows, int delta) {
    for (XSSFTable table : sheet.getTables()) {
      ExcelRange range = parsedRange(table.getCTTable().getRef(), "table", table.getName());
      if (affectsRows(range, rows, delta)) {
        throw unsupportedStructure(
            "SHIFT_ROWS",
            sheet,
            "table '" + table.getName() + "'",
            "row structural edits that would move tables are not supported");
      }
    }
    ExcelRange autofilter = sheetAutofilterRange(sheet);
    if (autofilter != null && affectsRows(autofilter, rows, delta)) {
      throw unsupportedStructure(
          "SHIFT_ROWS",
          sheet,
          "sheet autofilter " + ExcelSheetStructureSupport.formatRange(autofilter),
          "row structural edits that would move sheet autofilters are not supported");
    }
    for (String validationRange : dataValidationRanges(sheet)) {
      ExcelRange range = parsedRange(validationRange, "data validation", validationRange);
      if (affectsRows(range, rows, delta)) {
        throw unsupportedStructure(
            "SHIFT_ROWS",
            sheet,
            "data validation " + validationRange,
            "row structural edits that would move data validations are not supported");
      }
    }
    rejectDestructiveNamedRangesForRowShift(sheet.getWorkbook(), sheet, rows, delta);
  }

  static void rejectAffectedColumnStructuresForInsert(XSSFSheet sheet, int columnIndex) {
    for (XSSFTable table : sheet.getTables()) {
      ExcelRange range = parsedRange(table.getCTTable().getRef(), "table", table.getName());
      if (range.lastColumn() >= columnIndex) {
        throw unsupportedStructure(
            "INSERT_COLUMNS",
            sheet,
            "table '" + table.getName() + "'",
            "column structural edits that would move tables are not supported");
      }
    }
    ExcelRange autofilter = sheetAutofilterRange(sheet);
    if (autofilter != null && autofilter.lastColumn() >= columnIndex) {
      throw unsupportedStructure(
          "INSERT_COLUMNS",
          sheet,
          "sheet autofilter " + ExcelSheetStructureSupport.formatRange(autofilter),
          "column structural edits that would move sheet autofilters are not supported");
    }
  }

  static void rejectAffectedColumnStructuresForDelete(XSSFSheet sheet, ExcelColumnSpan columns) {
    for (XSSFTable table : sheet.getTables()) {
      ExcelRange range = parsedRange(table.getCTTable().getRef(), "table", table.getName());
      if (range.lastColumn() >= columns.firstColumnIndex()) {
        throw unsupportedStructure(
            "DELETE_COLUMNS",
            sheet,
            "table '" + table.getName() + "'",
            "column structural edits that would move or truncate tables are not supported");
      }
    }
    ExcelRange autofilter = sheetAutofilterRange(sheet);
    if (autofilter != null && autofilter.lastColumn() >= columns.firstColumnIndex()) {
      throw unsupportedStructure(
          "DELETE_COLUMNS",
          sheet,
          "sheet autofilter " + ExcelSheetStructureSupport.formatRange(autofilter),
          "column structural edits that would move or truncate sheet autofilters are not supported");
    }
    for (String validationRange : dataValidationRanges(sheet)) {
      ExcelRange range = parsedRange(validationRange, "data validation", validationRange);
      if (range.lastColumn() >= columns.firstColumnIndex()) {
        throw unsupportedStructure(
            "DELETE_COLUMNS",
            sheet,
            "data validation " + validationRange,
            "column structural edits that would move or truncate data validations are not supported");
      }
    }
    rejectDestructiveNamedRangesForColumnDelete(sheet.getWorkbook(), sheet, columns);
  }

  static void rejectAffectedColumnStructuresForShift(
      XSSFSheet sheet, ExcelColumnSpan columns, int delta) {
    for (XSSFTable table : sheet.getTables()) {
      ExcelRange range = parsedRange(table.getCTTable().getRef(), "table", table.getName());
      if (affectsColumns(range, columns, delta)) {
        throw unsupportedStructure(
            "SHIFT_COLUMNS",
            sheet,
            "table '" + table.getName() + "'",
            "column structural edits that would move tables are not supported");
      }
    }
    ExcelRange autofilter = sheetAutofilterRange(sheet);
    if (autofilter != null && affectsColumns(autofilter, columns, delta)) {
      throw unsupportedStructure(
          "SHIFT_COLUMNS",
          sheet,
          "sheet autofilter " + ExcelSheetStructureSupport.formatRange(autofilter),
          "column structural edits that would move sheet autofilters are not supported");
    }
    for (String validationRange : dataValidationRanges(sheet)) {
      ExcelRange range = parsedRange(validationRange, "data validation", validationRange);
      if (affectsColumns(range, columns, delta)) {
        throw unsupportedStructure(
            "SHIFT_COLUMNS",
            sheet,
            "data validation " + validationRange,
            "column structural edits that would move data validations are not supported");
      }
    }
    rejectDestructiveNamedRangesForColumnShift(sheet.getWorkbook(), sheet, columns, delta);
  }

  static void rejectDestructiveNamedRangesForRowDelete(
      XSSFWorkbook workbook, XSSFSheet sheet, ExcelRowSpan rows) {
    for (ResolvedNamedRange namedRange :
        resolvedRangeBackedNames(workbook, workbook.getAllNames())) {
      if (!namedRange.targetsSheet(sheet.getSheetName())) {
        continue;
      }
      if (overlapsRows(namedRange.range(), rows)) {
        throw unsupportedStructure(
            "DELETE_ROWS",
            sheet,
            namedRange.label(),
            "row structural edits that would truncate range-backed named ranges are not supported");
      }
    }
  }

  static void rejectDestructiveNamedRangesForRowShift(
      XSSFWorkbook workbook, XSSFSheet sheet, ExcelRowSpan rows, int delta) {
    for (ResolvedNamedRange namedRange :
        resolvedRangeBackedNames(workbook, workbook.getAllNames())) {
      if (!namedRange.targetsSheet(sheet.getSheetName())) {
        continue;
      }
      if (shiftWouldCorruptRows(namedRange.range(), rows, delta)) {
        throw unsupportedStructure(
            "SHIFT_ROWS",
            sheet,
            namedRange.label(),
            "row structural edits that would overwrite or partially move range-backed named ranges are not supported");
      }
    }
  }

  static void rejectDestructiveNamedRangesForColumnDelete(
      XSSFWorkbook workbook, XSSFSheet sheet, ExcelColumnSpan columns) {
    for (ResolvedNamedRange namedRange :
        resolvedRangeBackedNames(workbook, workbook.getAllNames())) {
      if (!namedRange.targetsSheet(sheet.getSheetName())) {
        continue;
      }
      if (overlapsColumns(namedRange.range(), columns)) {
        throw unsupportedStructure(
            "DELETE_COLUMNS",
            sheet,
            namedRange.label(),
            "column structural edits that would truncate range-backed named ranges are not supported");
      }
    }
  }

  static void rejectDestructiveNamedRangesForColumnShift(
      XSSFWorkbook workbook, XSSFSheet sheet, ExcelColumnSpan columns, int delta) {
    for (ResolvedNamedRange namedRange :
        resolvedRangeBackedNames(workbook, workbook.getAllNames())) {
      if (!namedRange.targetsSheet(sheet.getSheetName())) {
        continue;
      }
      if (shiftWouldCorruptColumns(namedRange.range(), columns, delta)) {
        throw unsupportedStructure(
            "SHIFT_COLUMNS",
            sheet,
            namedRange.label(),
            "column structural edits that would overwrite or partially move range-backed named ranges are not supported");
      }
    }
  }

  static void rejectFormulaBearingWorkbookForColumnEdit(
      XSSFWorkbook workbook, String operationType) {
    if (workbookContainsFormulas(workbook)) {
      throw new IllegalArgumentException(
          operationType
              + " cannot run while workbook formulas are present; Apache POI leaves some column references stale during column structural edits");
    }
    if (workbookContainsFormulaDefinedNames(workbook)) {
      throw new IllegalArgumentException(
          operationType
              + " cannot run while formula-defined names are present; Apache POI leaves some column references stale during column structural edits");
    }
  }

  private static boolean workbookContainsFormulas(XSSFWorkbook workbook) {
    for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
      for (Row row : workbook.getSheetAt(sheetIndex)) {
        for (Cell cell : row) {
          if (cell.getCellType() == CellType.FORMULA) {
            return true;
          }
        }
      }
    }
    return false;
  }

  static boolean workbookContainsFormulaDefinedNames(XSSFWorkbook workbook) {
    return workbookContainsFormulaDefinedNames(workbook, workbook.getAllNames());
  }

  static boolean workbookContainsFormulaDefinedNames(
      XSSFWorkbook workbook, Iterable<? extends Name> names) {
    for (Name name : names) {
      String refersToFormula = name.getRefersToFormula();
      if (refersToFormula == null || refersToFormula.isBlank()) {
        continue;
      }
      ExcelNamedRangeScope scope = scopeForName(workbook, name);
      if (ExcelNamedRangeTargets.resolveTarget(refersToFormula, scope).isEmpty()) {
        return true;
      }
    }
    return false;
  }

  static Optional<ExcelNamedRangeTarget> resolvedRangeBackedTarget(
      XSSFWorkbook workbook, Name name) {
    String refersToFormula = name.getRefersToFormula();
    if (refersToFormula == null || refersToFormula.isBlank()) {
      return Optional.empty();
    }
    return ExcelNamedRangeTargets.resolveTarget(refersToFormula, scopeForName(workbook, name));
  }

  static List<ResolvedNamedRange> resolvedRangeBackedNames(
      XSSFWorkbook workbook, Iterable<? extends Name> names) {
    List<ResolvedNamedRange> resolved = new ArrayList<>();
    for (Name name : names) {
      resolvedRangeBackedTarget(workbook, name)
          .ifPresent(
              target ->
                  resolved.add(
                      new ResolvedNamedRange(
                          name.getNameName(), target, ExcelRange.parse(target.range()))));
    }
    return List.copyOf(resolved);
  }

  private static ExcelNamedRangeScope scopeForName(XSSFWorkbook workbook, Name name) {
    if (name.getSheetIndex() < 0) {
      return new ExcelNamedRangeScope.WorkbookScope();
    }
    return new ExcelNamedRangeScope.SheetScope(workbook.getSheetName(name.getSheetIndex()));
  }

  private static IllegalArgumentException unsupportedStructure(
      String operationType, XSSFSheet sheet, String structureLabel, String reason) {
    return new IllegalArgumentException(
        operationType
            + " cannot move "
            + structureLabel
            + " on sheet '"
            + sheet.getSheetName()
            + "'; "
            + reason);
  }

  static boolean affectsRows(ExcelRange range, ExcelRowSpan rows, int delta) {
    int affectedFirstRow = Math.min(rows.firstRowIndex(), rows.firstRowIndex() + delta);
    int affectedLastRow = Math.max(rows.lastRowIndex(), rows.lastRowIndex() + delta);
    return range.firstRow() <= affectedLastRow && range.lastRow() >= affectedFirstRow;
  }

  static boolean affectsColumns(ExcelRange range, ExcelColumnSpan columns, int delta) {
    int affectedFirstColumn =
        Math.min(columns.firstColumnIndex(), columns.firstColumnIndex() + delta);
    int affectedLastColumn = Math.max(columns.lastColumnIndex(), columns.lastColumnIndex() + delta);
    return range.firstColumn() <= affectedLastColumn && range.lastColumn() >= affectedFirstColumn;
  }

  private static boolean overlapsRows(ExcelRange range, ExcelRowSpan rows) {
    return range.firstRow() <= rows.lastRowIndex() && range.lastRow() >= rows.firstRowIndex();
  }

  private static boolean overlapsColumns(ExcelRange range, ExcelColumnSpan columns) {
    return range.firstColumn() <= columns.lastColumnIndex()
        && range.lastColumn() >= columns.firstColumnIndex();
  }

  private static boolean fullyWithinRows(ExcelRange range, ExcelRowSpan rows) {
    return range.firstRow() >= rows.firstRowIndex() && range.lastRow() <= rows.lastRowIndex();
  }

  private static boolean fullyWithinColumns(ExcelRange range, ExcelColumnSpan columns) {
    return range.firstColumn() >= columns.firstColumnIndex()
        && range.lastColumn() <= columns.lastColumnIndex();
  }

  static boolean shiftWouldCorruptRows(ExcelRange range, ExcelRowSpan rows, int delta) {
    if (fullyWithinRows(range, rows)) {
      return false;
    }
    ExcelRowSpan destination =
        new ExcelRowSpan(rows.firstRowIndex() + delta, rows.lastRowIndex() + delta);
    return overlapsRows(range, rows) || overlapsRows(range, destination);
  }

  static boolean shiftWouldCorruptColumns(ExcelRange range, ExcelColumnSpan columns, int delta) {
    if (fullyWithinColumns(range, columns)) {
      return false;
    }
    ExcelColumnSpan destination =
        new ExcelColumnSpan(columns.firstColumnIndex() + delta, columns.lastColumnIndex() + delta);
    return overlapsColumns(range, columns) || overlapsColumns(range, destination);
  }

  record ResolvedNamedRange(String name, ExcelNamedRangeTarget target, ExcelRange range) {
    ResolvedNamedRange {
      Objects.requireNonNull(name, "name must not be null");
      Objects.requireNonNull(target, "target must not be null");
      Objects.requireNonNull(range, "range must not be null");
    }

    private boolean targetsSheet(String sheetName) {
      return target.sheetName().equals(sheetName);
    }

    private String label() {
      return "named range '" + name + "'";
    }
  }

  private static ExcelRange sheetAutofilterRange(XSSFSheet sheet) {
    if (!sheet.getCTWorksheet().isSetAutoFilter()) {
      return null;
    }
    return parsedRange(
        sheet.getCTWorksheet().getAutoFilter().getRef(),
        "sheet autofilter",
        sheet.getCTWorksheet().getAutoFilter().getRef());
  }

  private static List<String> dataValidationRanges(XSSFSheet sheet) {
    CTDataValidations dataValidations = sheet.getCTWorksheet().getDataValidations();
    if (dataValidations == null) {
      return List.of();
    }
    List<String> ranges = new ArrayList<>();
    for (CTDataValidation validation : dataValidations.getDataValidationArray()) {
      ranges.addAll(ExcelSqrefSupport.normalizedSqref(validation.getSqref()));
    }
    return List.copyOf(ranges);
  }

  private static ExcelRange parsedRange(String rawRange, String structureType, String detail) {
    ExcelRange range = ExcelSheetStructureSupport.parseRangeOrNull(rawRange);
    if (range == null) {
      throw new IllegalArgumentException(
          "Stored " + structureType + " range is invalid and cannot be normalized: " + detail);
    }
    return range;
  }
}
