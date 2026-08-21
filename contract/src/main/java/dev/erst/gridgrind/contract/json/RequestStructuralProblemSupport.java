package dev.erst.gridgrind.contract.json;

import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Shared invariants for the small structural-problem value types. */
final class RequestStructuralProblemSupport {
  private RequestStructuralProblemSupport() {}

  static String requireText(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  static void requireByteOffset(long byteOffset) {
    if (byteOffset < 0) {
      throw new IllegalArgumentException("byteOffset must not be negative");
    }
  }

  static Optional<Long> copyByteOffset(Optional<Long> byteOffset) {
    Objects.requireNonNull(byteOffset, "byteOffset must not be null");
    byteOffset.ifPresent(RequestStructuralProblemSupport::requireByteOffset);
    return byteOffset;
  }

  static Optional<String> copyJsonPath(Optional<String> jsonPath) {
    Objects.requireNonNull(jsonPath, "jsonPath must not be null");
    jsonPath.ifPresent(path -> requireText(path, "jsonPath"));
    return jsonPath;
  }

  static List<String> copyStrings(List<String> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    return values.stream().map(value -> requireText(value, fieldName)).toList();
  }

  static Optional<String> copyOptionalText(Optional<String> value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    return value.map(text -> requireText(text, fieldName));
  }

  static Optional<String> optionalJsonPath(String jsonPath) {
    Objects.requireNonNull(jsonPath, "jsonPath must not be null");
    return jsonPath.isEmpty() ? Optional.empty() : Optional.of(jsonPath);
  }
}
