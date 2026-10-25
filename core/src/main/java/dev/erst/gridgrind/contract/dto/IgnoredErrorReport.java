package dev.erst.gridgrind.contract.dto;

import dev.erst.gridgrind.excel.foundation.ExcelIgnoredErrorType;
import java.util.ArrayList;
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
    errorTypes = copyValues(errorTypes, "errorTypes");
    if (errorTypes.isEmpty()) {
      throw new IllegalArgumentException("errorTypes must not be empty");
    }
    if (errorTypes.size() != new LinkedHashSet<>(errorTypes).size()) {
      throw new IllegalArgumentException("errorTypes must not contain duplicates");
    }
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
