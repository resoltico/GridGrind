package dev.erst.gridgrind.core.vocabulary;

import java.util.Objects;
import org.jspecify.annotations.Nullable;

/** One primitive text invariant shared by validated values and non-identifying text fields. */
public final class TextValidation {
  private TextValidation() {}

  /** Returns the exact text, rejecting null or Unicode-blank values with field-specific wording. */
  public static String requireNonBlank(@Nullable String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
