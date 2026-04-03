package dev.erst.gridgrind.protocol.dto;

import java.util.List;
import java.util.Objects;

/** Data-validation-health analysis for one selected sheet set. */
public record DataValidationHealthReport(
    int checkedValidationCount,
    GridGrindResponse.AnalysisSummaryReport summary,
    List<GridGrindResponse.AnalysisFindingReport> findings) {
  public DataValidationHealthReport {
    if (checkedValidationCount < 0) {
      throw new IllegalArgumentException("checkedValidationCount must not be negative");
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
