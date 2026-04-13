package dev.erst.gridgrind.excel;

import java.util.List;
import java.util.Objects;

/** Authoritative sheet-presentation state for display flags, tab color, defaults, and warnings. */
public record ExcelSheetPresentation(
    ExcelSheetDisplay display,
    ExcelColor tabColor,
    ExcelSheetOutlineSummary outlineSummary,
    ExcelSheetDefaults sheetDefaults,
    List<ExcelIgnoredError> ignoredErrors) {
  /** Returns the effective default sheet-presentation state for one new worksheet. */
  public static ExcelSheetPresentation defaults() {
    return new ExcelSheetPresentation(
        ExcelSheetDisplay.defaults(),
        null,
        ExcelSheetOutlineSummary.defaults(),
        ExcelSheetDefaults.defaults(),
        List.of());
  }

  public ExcelSheetPresentation {
    Objects.requireNonNull(display, "display must not be null");
    Objects.requireNonNull(outlineSummary, "outlineSummary must not be null");
    Objects.requireNonNull(sheetDefaults, "sheetDefaults must not be null");
    ignoredErrors =
        List.copyOf(Objects.requireNonNull(ignoredErrors, "ignoredErrors must not be null"));
  }
}
