package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;

/** Shared validation and copying helpers for split workbook read-result families. */
final class WorkbookResultSupport {
  private WorkbookResultSupport() {}

  static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  static List<String> copyStrings(List<String> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    List<String> copy = List.copyOf(values);
    for (String value : copy) {
      Objects.requireNonNull(value, fieldName + " must not contain nulls");
    }
    return copy;
  }

  static List<String> copyDistinctStrings(List<String> values, String fieldName) {
    List<String> copy = copyStrings(values, fieldName);
    if (copy.size() != new LinkedHashSet<>(copy).size()) {
      throw new IllegalArgumentException(fieldName + " must not contain duplicates");
    }
    return copy;
  }

  static List<String> validateCommonWorkbookSummaryFields(
      int sheetCount, List<String> sheetNames, int namedRangeCount) {
    if (sheetCount < 0) {
      throw new IllegalArgumentException("sheetCount must not be negative");
    }
    if (namedRangeCount < 0) {
      throw new IllegalArgumentException("namedRangeCount must not be negative");
    }
    List<String> copy = copyDistinctStrings(sheetNames, "sheetNames");
    if (sheetCount != copy.size()) {
      throw new IllegalArgumentException("sheetCount must match sheetNames size");
    }
    for (String sheetName : copy) {
      if (sheetName.isBlank()) {
        throw new IllegalArgumentException("sheetNames must not contain blank values");
      }
    }
    return copy;
  }

  static <T> List<T> copyValues(List<T> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    List<T> copy = new ArrayList<>(values.size());
    for (T value : values) {
      copy.add(Objects.requireNonNull(value, fieldName + " must not contain nulls"));
    }
    return List.copyOf(copy);
  }
}
