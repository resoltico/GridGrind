package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.layout.SheetDefaultsReport;
import dev.erst.gridgrind.core.model.layout.SheetDefaultsSpec;
import dev.erst.gridgrind.core.model.layout.SheetDisplayReport;
import dev.erst.gridgrind.core.model.layout.SheetDisplaySpec;
import dev.erst.gridgrind.core.model.layout.SheetOutlineSummaryReport;
import dev.erst.gridgrind.core.model.layout.SheetOutlineSummarySpec;
import dev.erst.gridgrind.core.model.layout.SheetPresentationReport;
import dev.erst.gridgrind.core.model.layout.SheetPresentationSpec;
import dev.erst.gridgrind.core.model.style.CellColorReport;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.model.validation.IgnoredErrorReport;
import dev.erst.gridgrind.core.model.validation.IgnoredErrorSpec;
import dev.erst.gridgrind.core.vocabulary.ExcelIgnoredErrorType;
import dev.erst.gridgrind.core.vocabulary.ExcelSheetLayoutLimits;
import java.util.Comparator;
import java.util.EnumSet;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;
import java.util.TreeMap;
import org.apache.poi.ss.usermodel.IgnoredErrorType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/** Applies and reads cohesive sheet-presentation state such as display flags and defaults. */
final class ExcelSheetPresentationController {
  /** Applies the provided presentation payload authoritatively to the target sheet. */
  void setPresentation(XSSFSheet sheet, SheetPresentationSpec presentation) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    Objects.requireNonNull(presentation, "presentation must not be null");

    applyDisplay(sheet, presentation.display());
    applyTabColor(sheet, presentation.tabColor());
    applyOutlineSummary(sheet, presentation.outlineSummary());
    applyDefaults(sheet, presentation.sheetDefaults());
    applyIgnoredErrors(sheet, presentation.ignoredErrors());
  }

  /** Clears sheet presentation back to the effective workbook defaults. */
  void clearPresentation(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    setPresentation(sheet, SheetPresentationSpec.defaults());
  }

  /** Returns the factual sheet-presentation state currently stored on the provided sheet. */
  SheetPresentationReport presentation(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    return new SheetPresentationReport(
        display(sheet),
        tabColor(sheet),
        outlineSummary(sheet),
        defaults(sheet),
        ignoredErrors(sheet));
  }

  private static void applyDisplay(XSSFSheet sheet, SheetDisplaySpec display) {
    sheet.setDisplayGridlines(display.displayGridlines());
    sheet.setDisplayZeros(display.displayZeros());
    sheet.setDisplayRowColHeadings(display.displayRowColHeadings());
    sheet.setDisplayFormulas(display.displayFormulas());
    sheet.setRightToLeft(display.rightToLeft());
  }

  private static SheetDisplayReport display(XSSFSheet sheet) {
    return new SheetDisplayReport(
        sheet.isDisplayGridlines(),
        sheet.isDisplayZeros(),
        sheet.isDisplayRowColHeadings(),
        sheet.isDisplayFormulas(),
        sheet.isRightToLeft());
  }

  private static void applyTabColor(XSSFSheet sheet, Optional<ColorSpec> tabColor) {
    Objects.requireNonNull(tabColor, "tabColor must not be null");
    if (tabColor.isEmpty()) {
      clearTabColor(sheet);
      return;
    }
    sheet.setTabColor(ExcelColorSupport.toXssfColor(sheet.getWorkbook(), tabColor.orElseThrow()));
  }

  private static Optional<CellColorReport> tabColor(XSSFSheet sheet) {
    XSSFColor color = sheet.getTabColor();
    return ExcelColorSnapshotSupport.snapshot(color);
  }

  static void clearTabColor(XSSFSheet sheet) {
    if (!sheet.getCTWorksheet().isSetSheetPr()) {
      return;
    }
    var sheetPr = sheet.getCTWorksheet().getSheetPr();
    if (sheetPr.isSetTabColor()) {
      sheetPr.unsetTabColor();
    }
  }

  private static void applyOutlineSummary(XSSFSheet sheet, SheetOutlineSummarySpec outlineSummary) {
    sheet.setRowSumsBelow(outlineSummary.rowSumsBelow());
    sheet.setRowSumsRight(outlineSummary.rowSumsRight());
  }

  private static SheetOutlineSummaryReport outlineSummary(XSSFSheet sheet) {
    return new SheetOutlineSummaryReport(sheet.getRowSumsBelow(), sheet.getRowSumsRight());
  }

  private static void applyDefaults(XSSFSheet sheet, SheetDefaultsSpec defaults) {
    ExcelSheetLayoutLimits.requireDefaultColumnWidth(
        defaults.defaultColumnWidth(), "defaultColumnWidth");
    ExcelSheetLayoutLimits.requireRowHeightPoints(
        defaults.defaultRowHeightPoints(), "defaultRowHeightPoints");
    sheet.setDefaultColumnWidth(defaults.defaultColumnWidth());
    sheet.setDefaultRowHeightInPoints((float) defaults.defaultRowHeightPoints());
  }

  private static SheetDefaultsReport defaults(XSSFSheet sheet) {
    double defaultRowHeightPoints = sheet.getDefaultRowHeightInPoints();
    return new SheetDefaultsReport(
        sheet.getDefaultColumnWidth(),
        defaultRowHeightPoints <= 0.0d
            ? SheetDefaultsReport.defaults().defaultRowHeightPoints()
            : defaultRowHeightPoints);
  }

  private static void applyIgnoredErrors(XSSFSheet sheet, List<IgnoredErrorSpec> ignoredErrors) {
    clearIgnoredErrors(sheet);
    for (IgnoredErrorSpec ignoredError : ignoredErrors) {
      sheet.addIgnoredErrors(
          ExcelSheetStructureSupport.toCellRangeAddress(
              ExcelRange.parse(ignoredError.range().value())),
          ignoredError.errorTypes().stream()
              .map(ExcelIgnoredErrorPoiBridge::toPoi)
              .toArray(IgnoredErrorType[]::new));
    }
  }

  private static void clearIgnoredErrors(XSSFSheet sheet) {
    if (sheet.getCTWorksheet().isSetIgnoredErrors()) {
      sheet.getCTWorksheet().unsetIgnoredErrors();
    }
  }

  @SuppressWarnings({"PMD.LooseCoupling", "PMD.UseConcurrentHashMap"})
  private static List<IgnoredErrorReport> ignoredErrors(XSSFSheet sheet) {
    Map<String, EnumSet<ExcelIgnoredErrorType>> ignoredErrorsByRange = new TreeMap<>();
    for (Map.Entry<IgnoredErrorType, java.util.Set<CellRangeAddress>> entry :
        sheet.getIgnoredErrors().entrySet()) {
      ExcelIgnoredErrorType errorType = ExcelIgnoredErrorPoiBridge.fromPoi(entry.getKey());
      entry.getValue().stream()
          .sorted(Comparator.comparing(CellRangeAddress::formatAsString))
          .map(CellRangeAddress::formatAsString)
          .forEach(
              range ->
                  ignoredErrorsByRange
                      .computeIfAbsent(
                          range, ignoredError -> EnumSet.noneOf(ExcelIgnoredErrorType.class))
                      .add(errorType));
    }
    return ignoredErrorsByRange.entrySet().stream()
        .map(
            entry ->
                new IgnoredErrorReport(
                    new dev.erst.gridgrind.core.vocabulary.RangeLabel(entry.getKey()),
                    List.copyOf(entry.getValue())))
        .toList();
  }
}
