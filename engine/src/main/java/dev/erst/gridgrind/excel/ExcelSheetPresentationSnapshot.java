package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.style.CellColorReport;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Factual sheet-presentation state loaded from one worksheet. */
public record ExcelSheetPresentationSnapshot(
    ExcelSheetDisplay display,
    Optional<CellColorReport> tabColor,
    ExcelSheetOutlineSummary outlineSummary,
    ExcelSheetDefaults sheetDefaults,
    List<ExcelIgnoredError> ignoredErrors) {
  public ExcelSheetPresentationSnapshot {
    Objects.requireNonNull(display, "display must not be null");
    Objects.requireNonNull(tabColor, "tabColor must not be null");
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
        display,
        tabColor.map(ExcelColorSupport::copyOf),
        outlineSummary,
        sheetDefaults,
        ignoredErrors);
  }
}
