package dev.erst.gridgrind.excel.drawing;

import java.util.Objects;

/** Drawing-package argument validation helpers. */
public final class ExcelDrawingArgumentSupport {
  private ExcelDrawingArgumentSupport() {}

  /** Returns the value after enforcing that it is present and not blank. */
  public static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
