package dev.erst.gridgrind.protocol;

import java.util.List;
import java.util.Objects;

/** Table-health analysis for one selected table set. */
public record TableHealthReport(
    int checkedTableCount,
    GridGrindResponse.AnalysisSummaryReport summary,
    List<GridGrindResponse.AnalysisFindingReport> findings) {
  public TableHealthReport {
    if (checkedTableCount < 0) {
      throw new IllegalArgumentException("checkedTableCount must not be negative");
    }
    Objects.requireNonNull(summary, "summary must not be null");
    findings = copyValues(findings, "findings");
  }

  private static <T> List<T> copyValues(List<T> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    List<T> copy = List.copyOf(values);
    for (T value : copy) {
      Objects.requireNonNull(value, fieldName + " must not contain nulls");
    }
    return copy;
  }
}
