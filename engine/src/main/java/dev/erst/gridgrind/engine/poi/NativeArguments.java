package dev.erst.gridgrind.engine.poi;

import java.util.Objects;

/** Shared argument validation helpers for workbook-core seams. */
public final class NativeArguments {
  private NativeArguments() {}

  /** Checks the native string boundary while retaining the field-specific failure message. */
  public static void requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
  }
}
