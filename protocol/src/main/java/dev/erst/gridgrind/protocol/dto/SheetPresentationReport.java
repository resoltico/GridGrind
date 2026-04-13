package dev.erst.gridgrind.protocol.dto;

import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;

/** Factual sheet-presentation state reported for one worksheet. */
public record SheetPresentationReport(
    SheetDisplayReport display,
    CellColorReport tabColor,
    SheetOutlineSummaryReport outlineSummary,
    SheetDefaultsReport sheetDefaults,
    List<IgnoredErrorReport> ignoredErrors) {
  /** Returns the effective default sheet-presentation state for one new worksheet. */
  public static SheetPresentationReport defaults() {
    return new SheetPresentationReport(
        SheetDisplayReport.defaults(),
        null,
        SheetOutlineSummaryReport.defaults(),
        SheetDefaultsReport.defaults(),
        List.of());
  }

  public SheetPresentationReport {
    Objects.requireNonNull(display, "display must not be null");
    Objects.requireNonNull(outlineSummary, "outlineSummary must not be null");
    Objects.requireNonNull(sheetDefaults, "sheetDefaults must not be null");
    ignoredErrors =
        List.copyOf(Objects.requireNonNull(ignoredErrors, "ignoredErrors must not be null"));
    List<String> normalizedRanges = ignoredErrors.stream().map(IgnoredErrorReport::range).toList();
    if (normalizedRanges.size() != new LinkedHashSet<>(normalizedRanges).size()) {
      throw new IllegalArgumentException("ignoredErrors must not contain duplicate ranges");
    }
  }
}
