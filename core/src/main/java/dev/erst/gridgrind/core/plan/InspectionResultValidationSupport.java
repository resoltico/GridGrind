package dev.erst.gridgrind.core.plan;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Shared validation for inspection result records. */
final class InspectionResultValidationSupport {
  private InspectionResultValidationSupport() {}

  static <T> List<T> copyValues(List<T> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    List<T> copy = new ArrayList<>(values.size());
    for (T value : values) {
      copy.add(Objects.requireNonNull(value, fieldName + " must not contain nulls"));
    }
    return List.copyOf(copy);
  }
}
