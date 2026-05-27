package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Shared argument validation helpers for workbook-core seams. */
final class ExcelArgumentSupport {
  private ExcelArgumentSupport() {}

  static void requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
  }
}
