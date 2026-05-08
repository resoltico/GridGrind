package dev.erst.gridgrind.contract.dto;

import java.util.List;

/** High-level characterization of the named ranges selected by one analysis read. */
public record NamedRangeSurfaceReport(
    int workbookScopedCount,
    int sheetScopedCount,
    int rangeBackedCount,
    int formulaBackedCount,
    List<NamedRangeSurfaceEntryReport> namedRanges) {
  public NamedRangeSurfaceReport {
    if (workbookScopedCount < 0) {
      throw new IllegalArgumentException("workbookScopedCount must not be negative");
    }
    if (sheetScopedCount < 0) {
      throw new IllegalArgumentException("sheetScopedCount must not be negative");
    }
    if (rangeBackedCount < 0) {
      throw new IllegalArgumentException("rangeBackedCount must not be negative");
    }
    if (formulaBackedCount < 0) {
      throw new IllegalArgumentException("formulaBackedCount must not be negative");
    }
    namedRanges = GridGrindResponseSupport.copyValues(namedRanges, "namedRanges");
  }
}
