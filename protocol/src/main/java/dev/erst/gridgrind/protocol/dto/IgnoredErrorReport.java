package dev.erst.gridgrind.protocol.dto;

import dev.erst.gridgrind.excel.ExcelIgnoredErrorType;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;

/** Factual ignored-error block reported for one A1-style range. */
public record IgnoredErrorReport(String range, List<ExcelIgnoredErrorType> errorTypes) {
  public IgnoredErrorReport {
    Objects.requireNonNull(range, "range must not be null");
    if (range.isBlank()) {
      throw new IllegalArgumentException("range must not be blank");
    }
    errorTypes = List.copyOf(Objects.requireNonNull(errorTypes, "errorTypes must not be null"));
    if (errorTypes.isEmpty()) {
      throw new IllegalArgumentException("errorTypes must not be empty");
    }
    if (errorTypes.size() != new LinkedHashSet<>(errorTypes).size()) {
      throw new IllegalArgumentException("errorTypes must not contain duplicates");
    }
  }
}
