package dev.erst.gridgrind.contract.dto;

import java.util.ArrayList;
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
    List<T> copy = new ArrayList<>(values.size());
    for (T value : values) {
      copy.add(Objects.requireNonNull(value, fieldName + " must not contain nulls"));
    }
    return List.copyOf(copy);
  }
}
