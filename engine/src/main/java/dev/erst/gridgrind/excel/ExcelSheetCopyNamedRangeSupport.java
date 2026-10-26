package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.formula.NamedRangeReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeScopeReport;
import dev.erst.gridgrind.core.model.formula.NamedRangeTarget;
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
      List<NamedRangeReport> localNamedRanges) {
    deleteLocalNamedRanges(workbook, targetSheetName);
    copyLocalNamedRanges(workbook, sourceSheetName, targetSheetName, localNamedRanges);
  }

  static void deleteLocalNamedRanges(ExcelWorkbook workbook, String sheetName) {
    NamedRangeScopeReport.Sheet scope =
        new NamedRangeScopeReport.Sheet(
            new dev.erst.gridgrind.core.vocabulary.SheetLabel(sheetName));
    for (NamedRangeReport localName :
        copyableLocalNames(workbook.names().namedRanges(), sheetName)) {
      workbook.names().deleteNamedRange(localName.name().value(), scope);
    }
  }

  static List<NamedRangeReport> copyableLocalNames(
      List<NamedRangeReport> namedRanges, String sheetName) {
    Objects.requireNonNull(namedRanges, "namedRanges must not be null");
    ExcelWorkbookSheetSupport.requireSheetName(sheetName, "sheetName");
    List<NamedRangeReport> localNamedRanges = new ArrayList<>();
    for (NamedRangeReport namedRange : namedRanges) {
      switch (namedRange.scope()) {
        case NamedRangeScopeReport.Workbook _ -> {}
        case NamedRangeScopeReport.Sheet sheetScope -> {
          if (sheetScope.sheetName().value().equals(sheetName)) {
            localNamedRanges.add(namedRange);
          }
        }
      }
    }
    return List.copyOf(localNamedRanges);
  }

  private static void copyLocalNamedRanges(
      ExcelWorkbook workbook,
      String sourceSheetName,
      String targetSheetName,
      List<NamedRangeReport> localNamedRanges) {
    int targetSheetIndex = workbook.xssfWorkbook().getSheetIndex(targetSheetName);
    var destination = new dev.erst.gridgrind.core.vocabulary.SheetName(targetSheetName);
    for (NamedRangeReport localName : localNamedRanges) {
      String formula =
          switch (localName) {
            case NamedRangeReport.RangeReport range ->
                ExcelNamedRangeTargets.rangeFormula(
                    destination, ((NamedRangeTarget.Range) range.target()).range());
            case NamedRangeReport.FormulaReport value ->
                value.refersToFormula().isEmpty()
                    ? ""
                    : ExcelSheetCopySupport.retargetFormula(
                        workbook,
                        value.refersToFormula(),
                        FormulaType.NAMEDRANGE,
                        targetSheetIndex,
                        sourceSheetName,
                        targetSheetName);
          };
      var copied = workbook.xssfWorkbook().createName();
      copied.setNameName(localName.name().value());
      copied.setSheetIndex(targetSheetIndex);
      if (!formula.isEmpty()) copied.setRefersToFormula(formula);
    }
  }
}
