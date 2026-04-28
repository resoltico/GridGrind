package dev.erst.gridgrind.excel;

import java.util.List;
import java.util.Objects;

/** Factual sheet-presentation state loaded from one worksheet. */
public record ExcelSheetPresentationSnapshot(
    ExcelSheetDisplay display,
    ExcelColorSnapshot tabColor,
    ExcelSheetOutlineSummary outlineSummary,
    ExcelSheetDefaults sheetDefaults,
    List<ExcelIgnoredError> ignoredErrors) {
  public ExcelSheetPresentationSnapshot {
    Objects.requireNonNull(display, "display must not be null");
    Objects.requireNonNull(outlineSummary, "outlineSummary must not be null");
    Objects.requireNonNull(sheetDefaults, "sheetDefaults must not be null");
    ignoredErrors =
        List.copyOf(Objects.requireNonNull(ignoredErrors, "ignoredErrors must not be null"));
  }

  /**
   * Converts factual sheet-presentation state into the matching authoritative authoring payload.
   */
  public ExcelSheetPresentation toAuthoringPresentation() {
    return new ExcelSheetPresentation(
        display, ExcelColorSupport.copyOf(tabColor), outlineSummary, sheetDefaults, ignoredErrors);
  }
}
