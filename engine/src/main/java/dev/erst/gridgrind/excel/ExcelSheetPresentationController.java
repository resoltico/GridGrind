package dev.erst.gridgrind.excel;

import java.util.Comparator;
import java.util.EnumSet;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.TreeMap;
import org.apache.poi.ss.usermodel.IgnoredErrorType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/** Applies and reads cohesive sheet-presentation state such as display flags and defaults. */
final class ExcelSheetPresentationController {
  /** Applies the provided presentation payload authoritatively to the target sheet. */
  void setPresentation(XSSFSheet sheet, ExcelSheetPresentation presentation) {
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
    setPresentation(sheet, ExcelSheetPresentation.defaults());
  }

  /** Returns the factual sheet-presentation state currently stored on the provided sheet. */
  ExcelSheetPresentationSnapshot presentation(XSSFSheet sheet) {
    Objects.requireNonNull(sheet, "sheet must not be null");
    return new ExcelSheetPresentationSnapshot(
        display(sheet),
        tabColor(sheet),
        outlineSummary(sheet),
        defaults(sheet),
        ignoredErrors(sheet));
  }

  private static void applyDisplay(XSSFSheet sheet, ExcelSheetDisplay display) {
    sheet.setDisplayGridlines(display.displayGridlines());
    sheet.setDisplayZeros(display.displayZeros());
    sheet.setDisplayRowColHeadings(display.displayRowColHeadings());
    sheet.setDisplayFormulas(display.displayFormulas());
    sheet.setRightToLeft(display.rightToLeft());
  }

  private static ExcelSheetDisplay display(XSSFSheet sheet) {
    return new ExcelSheetDisplay(
        sheet.isDisplayGridlines(),
        sheet.isDisplayZeros(),
        sheet.isDisplayRowColHeadings(),
        sheet.isDisplayFormulas(),
        sheet.isRightToLeft());
  }

  private static void applyTabColor(XSSFSheet sheet, ExcelColor tabColor) {
    if (tabColor == null) {
      clearTabColor(sheet);
      return;
    }
    sheet.setTabColor(ExcelColorSupport.toXssfColor(sheet.getWorkbook(), tabColor));
  }

  private static ExcelColorSnapshot tabColor(XSSFSheet sheet) {
    XSSFColor color = sheet.getTabColor();
    return color == null ? null : ExcelColorSnapshotSupport.snapshot(color);
  }

  private static void clearTabColor(XSSFSheet sheet) {
    if (!sheet.getCTWorksheet().isSetSheetPr()) {
      return;
    }
    var sheetPr = sheet.getCTWorksheet().getSheetPr();
    if (sheetPr.isSetTabColor()) {
      sheetPr.unsetTabColor();
    }
  }

  private static void applyOutlineSummary(
      XSSFSheet sheet, ExcelSheetOutlineSummary outlineSummary) {
    sheet.setRowSumsBelow(outlineSummary.rowSumsBelow());
    sheet.setRowSumsRight(outlineSummary.rowSumsRight());
  }

  private static ExcelSheetOutlineSummary outlineSummary(XSSFSheet sheet) {
    return new ExcelSheetOutlineSummary(sheet.getRowSumsBelow(), sheet.getRowSumsRight());
  }

  private static void applyDefaults(XSSFSheet sheet, ExcelSheetDefaults defaults) {
    sheet.setDefaultColumnWidth(defaults.defaultColumnWidth());
    sheet.setDefaultRowHeightInPoints((float) defaults.defaultRowHeightPoints());
  }

  private static ExcelSheetDefaults defaults(XSSFSheet sheet) {
    double defaultRowHeightPoints = sheet.getDefaultRowHeightInPoints();
    return new ExcelSheetDefaults(
        sheet.getDefaultColumnWidth(),
        defaultRowHeightPoints <= 0.0d
            ? ExcelSheetDefaults.defaults().defaultRowHeightPoints()
            : defaultRowHeightPoints);
  }

  private static void applyIgnoredErrors(XSSFSheet sheet, List<ExcelIgnoredError> ignoredErrors) {
    clearIgnoredErrors(sheet);
    for (ExcelIgnoredError ignoredError : ignoredErrors) {
      sheet.addIgnoredErrors(
          CellRangeAddress.valueOf(ignoredError.range()),
          ignoredError.errorTypes().stream()
              .map(ExcelIgnoredErrorType::toPoi)
              .toArray(IgnoredErrorType[]::new));
    }
  }

  private static void clearIgnoredErrors(XSSFSheet sheet) {
    if (sheet.getCTWorksheet().isSetIgnoredErrors()) {
      sheet.getCTWorksheet().unsetIgnoredErrors();
    }
  }

  @SuppressWarnings({"PMD.LooseCoupling", "PMD.UseConcurrentHashMap"})
  private static List<ExcelIgnoredError> ignoredErrors(XSSFSheet sheet) {
    Map<String, EnumSet<ExcelIgnoredErrorType>> ignoredErrorsByRange = new TreeMap<>();
    for (Map.Entry<IgnoredErrorType, java.util.Set<CellRangeAddress>> entry :
        sheet.getIgnoredErrors().entrySet()) {
      ExcelIgnoredErrorType errorType = ExcelIgnoredErrorType.fromPoi(entry.getKey());
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
        .map(entry -> new ExcelIgnoredError(entry.getKey(), List.copyOf(entry.getValue())))
        .toList();
  }
}
