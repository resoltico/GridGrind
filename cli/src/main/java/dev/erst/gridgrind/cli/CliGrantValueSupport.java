package dev.erst.gridgrind.cli;

import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Shared validation and path normalization for CLI-owned host-grant transport values. */
final class CliGrantValueSupport {
  private CliGrantValueSupport() {}

  static Path resolve(String path, Path base) {
    Path candidate = Path.of(path);
    return (candidate.isAbsolute() ? candidate : base.resolve(candidate))
        .toAbsolutePath()
        .normalize();
  }

  static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  static List<String> copyStrings(List<String> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    List<String> copied = new ArrayList<>(values.size());
    for (String value : values) {
      copied.add(requireNonBlank(value, fieldName));
    }
    return List.copyOf(copied);
  }

  static <T> List<T> copyValues(List<T> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    List<T> copied = new ArrayList<>(values.size());
    for (T value : values) {
      copied.add(Objects.requireNonNull(value, fieldName + " must not contain null values"));
    }
    return List.copyOf(copied);
  }
}
