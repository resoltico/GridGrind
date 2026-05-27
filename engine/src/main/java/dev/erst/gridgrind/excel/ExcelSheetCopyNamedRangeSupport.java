package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import org.apache.poi.ss.formula.FormulaType;

/** Copies and retargets sheet-scoped named ranges during sheet duplication flows. */
final class ExcelSheetCopyNamedRangeSupport {
  private ExcelSheetCopyNamedRangeSupport() {}

  static void replaceLocalNamedRanges(
      ExcelWorkbook workbook,
      String sourceSheetName,
      String targetSheetName,
      List<ExcelNamedRangeSnapshot> localNamedRanges) {
    deleteLocalNamedRanges(workbook, targetSheetName);
    copyLocalNamedRanges(workbook, sourceSheetName, targetSheetName, localNamedRanges);
  }

  static void deleteLocalNamedRanges(ExcelWorkbook workbook, String sheetName) {
    ExcelNamedRangeScope.SheetScope scope = new ExcelNamedRangeScope.SheetScope(sheetName);
    for (ExcelNamedRangeSnapshot localName :
        copyableLocalNames(workbook.names().namedRanges(), sheetName)) {
      workbook.names().deleteNamedRange(localName.name(), scope);
    }
  }

  static List<ExcelNamedRangeSnapshot> copyableLocalNames(
      List<ExcelNamedRangeSnapshot> namedRanges, String sheetName) {
    Objects.requireNonNull(namedRanges, "namedRanges must not be null");
    ExcelWorkbookSheetSupport.requireSheetName(sheetName, "sheetName");
    List<ExcelNamedRangeSnapshot> localNamedRanges = new ArrayList<>();
    for (ExcelNamedRangeSnapshot namedRange : namedRanges) {
      switch (namedRange.scope()) {
        case ExcelNamedRangeScope.WorkbookScope _ -> {}
        case ExcelNamedRangeScope.SheetScope sheetScope -> {
          if (sheetScope.sheetName().equals(sheetName)) {
            localNamedRanges.add(namedRange);
          }
        }
      }
    }
    return List.copyOf(localNamedRanges);
  }

  static List<ExcelNamedRangeSnapshot.RangeSnapshot> copyableLocalRangeNames(
      List<ExcelNamedRangeSnapshot> namedRanges, String sheetName) {
    return copyableLocalNames(namedRanges, sheetName).stream()
        .flatMap(
            namedRange ->
                namedRange instanceof ExcelNamedRangeSnapshot.RangeSnapshot rangeSnapshot
                    ? java.util.stream.Stream.of(rangeSnapshot)
                    : java.util.stream.Stream.empty())
        .toList();
  }

  static void requireNoUncopyableLocalNamedRanges(
      List<ExcelNamedRangeSnapshot> namedRanges, String sheetName) {
    Objects.requireNonNull(namedRanges, "namedRanges must not be null");
    ExcelWorkbookSheetSupport.requireSheetName(sheetName, "sheetName");
  }

  private static void copyLocalNamedRanges(
      ExcelWorkbook workbook,
      String sourceSheetName,
      String targetSheetName,
      List<ExcelNamedRangeSnapshot> localNamedRanges) {
    ExcelNamedRangeScope.SheetScope scope = new ExcelNamedRangeScope.SheetScope(targetSheetName);
    int targetSheetIndex = workbook.xssfWorkbook().getSheetIndex(targetSheetName);
    for (ExcelNamedRangeSnapshot localNamedRange : localNamedRanges) {
      workbook
          .names()
          .setNamedRange(
              copiedLocalNamedRange(
                  workbook,
                  localNamedRange,
                  scope,
                  targetSheetIndex,
                  sourceSheetName,
                  targetSheetName));
    }
  }

  private static ExcelNamedRangeDefinition copiedLocalNamedRange(
      ExcelWorkbook workbook,
      ExcelNamedRangeSnapshot localNamedRange,
      ExcelNamedRangeScope.SheetScope scope,
      int targetSheetIndex,
      String sourceSheetName,
      String targetSheetName) {
    return switch (localNamedRange) {
      case ExcelNamedRangeSnapshot.RangeSnapshot rangeSnapshot ->
          new ExcelNamedRangeDefinition(
              rangeSnapshot.name(),
              scope,
              ExcelNamedRangeTarget.range(
                  targetSheetName, ((ExcelNamedRangeTarget.Range) rangeSnapshot.target()).range()));
      case ExcelNamedRangeSnapshot.FormulaSnapshot formulaSnapshot ->
          new ExcelNamedRangeDefinition(
              formulaSnapshot.name(),
              scope,
              ExcelNamedRangeTarget.formula(
                  ExcelSheetCopySupport.retargetFormula(
                      workbook,
                      formulaSnapshot.refersToFormula(),
                      FormulaType.NAMEDRANGE,
                      targetSheetIndex,
                      sourceSheetName,
                      targetSheetName)));
    };
  }
}
