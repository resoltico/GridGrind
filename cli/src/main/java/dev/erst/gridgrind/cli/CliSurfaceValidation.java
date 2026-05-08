package dev.erst.gridgrind.cli;

import java.util.List;
import java.util.Objects;

/** Validation helpers for CLI-owned discovery metadata. */
final class CliSurfaceValidation {
  private CliSurfaceValidation() {}

  static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  static List<String> copyStrings(List<String> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    List<String> copy = new java.util.ArrayList<>(values.size());
    for (String value : values) {
      copy.add(requireNonBlank(value, fieldName));
    }
    return List.copyOf(copy);
  }
}
