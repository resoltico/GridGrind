package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelColumnSpan;
import dev.erst.gridgrind.excel.foundation.ExcelRowSpan;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** Shared structural guard and normalization helpers used by row and column controllers. */
final class ExcelRowColumnStructureController {
  ExcelRowColumnStructureController() {}

  void rejectAffectedRowStructuresForInsert(XSSFSheet sheet, int rowIndex) {
    ExcelRowColumnStructureGuardSupport.rejectAffectedRowStructuresForInsert(sheet, rowIndex);
  }

  void rejectAffectedRowStructuresForDelete(XSSFSheet sheet, ExcelRowSpan rows) {
    ExcelRowColumnStructureGuardSupport.rejectAffectedRowStructuresForDelete(sheet, rows);
  }

  void rejectAffectedRowStructuresForShift(XSSFSheet sheet, ExcelRowSpan rows, int delta) {
    ExcelRowColumnStructureGuardSupport.rejectAffectedRowStructuresForShift(sheet, rows, delta);
  }

  void rejectAffectedColumnStructuresForInsert(XSSFSheet sheet, int columnIndex) {
    ExcelRowColumnStructureGuardSupport.rejectAffectedColumnStructuresForInsert(sheet, columnIndex);
  }

  void rejectAffectedColumnStructuresForDelete(XSSFSheet sheet, ExcelColumnSpan columns) {
    ExcelRowColumnStructureGuardSupport.rejectAffectedColumnStructuresForDelete(sheet, columns);
  }

  void rejectAffectedColumnStructuresForShift(XSSFSheet sheet, ExcelColumnSpan columns, int delta) {
    ExcelRowColumnStructureGuardSupport.rejectAffectedColumnStructuresForShift(
        sheet, columns, delta);
  }

  void rejectDestructiveNamedRangesForRowDelete( // LIM-018
      XSSFWorkbook workbook, XSSFSheet sheet, ExcelRowSpan rows) {
    ExcelRowColumnStructureGuardSupport.rejectDestructiveNamedRangesForRowDelete(
        workbook, sheet, rows);
  }

  void rejectDestructiveNamedRangesForRowShift( // LIM-018
      XSSFWorkbook workbook, XSSFSheet sheet, ExcelRowSpan rows, int delta) {
    ExcelRowColumnStructureGuardSupport.rejectDestructiveNamedRangesForRowShift(
        workbook, sheet, rows, delta);
  }

  void rejectDestructiveNamedRangesForColumnDelete( // LIM-018
      XSSFWorkbook workbook, XSSFSheet sheet, ExcelColumnSpan columns) {
    ExcelRowColumnStructureGuardSupport.rejectDestructiveNamedRangesForColumnDelete(
        workbook, sheet, columns);
  }

  void rejectDestructiveNamedRangesForColumnShift( // LIM-018
      XSSFWorkbook workbook, XSSFSheet sheet, ExcelColumnSpan columns, int delta) {
    ExcelRowColumnStructureGuardSupport.rejectDestructiveNamedRangesForColumnShift(
        workbook, sheet, columns, delta);
  }

  static boolean workbookContainsFormulaDefinedNames(
      XSSFWorkbook workbook, Iterable<? extends Name> names) {
    return ExcelRowColumnStructureGuardSupport.workbookContainsFormulaDefinedNames(workbook, names);
  }

  static Optional<ExcelNamedRangeTarget> resolvedRangeBackedTarget(
      XSSFWorkbook workbook, Name name) {
    return ExcelRowColumnStructureGuardSupport.resolvedRangeBackedTarget(workbook, name);
  }

  static List<ResolvedNamedRange> resolvedRangeBackedNames(
      XSSFWorkbook workbook, Iterable<? extends Name> names) {
    return ExcelRowColumnStructureGuardSupport.resolvedRangeBackedNames(workbook, names).stream()
        .map(result -> new ResolvedNamedRange(result.name(), result.target(), result.range()))
        .toList();
  }

  static boolean affectsRows(ExcelRange range, ExcelRowSpan rows, int delta) {
    return ExcelRowColumnStructureGuardSupport.affectsRows(range, rows, delta);
  }

  static boolean affectsColumns(ExcelRange range, ExcelColumnSpan columns, int delta) {
    return ExcelRowColumnStructureGuardSupport.affectsColumns(range, columns, delta);
  }

  static boolean shiftWouldCorruptRows(ExcelRange range, ExcelRowSpan rows, int delta) {
    return ExcelRowColumnStructureGuardSupport.shiftWouldCorruptRows(range, rows, delta);
  }

  static boolean shiftWouldCorruptColumns(ExcelRange range, ExcelColumnSpan columns, int delta) {
    return ExcelRowColumnStructureGuardSupport.shiftWouldCorruptColumns(range, columns, delta);
  }

  static void setColumnCollapsed(XSSFSheet sheet, int columnIndex, boolean collapsed) {
    ExcelRowColumnOutlineSupport.setColumnCollapsed(sheet, columnIndex, collapsed);
  }

  static void canonicalizeColumnDefinitions(XSSFSheet sheet) {
    ExcelRowColumnOutlineSupport.canonicalizeColumnDefinitions(sheet);
  }

  /** Typed resolved view of a range-backed defined name for structural guard evaluation. */
  record ResolvedNamedRange(String name, ExcelNamedRangeTarget target, ExcelRange range) {
    ResolvedNamedRange {
      Objects.requireNonNull(name, "name must not be null");
      Objects.requireNonNull(target, "target must not be null");
      Objects.requireNonNull(range, "range must not be null");
    }
  }
}
