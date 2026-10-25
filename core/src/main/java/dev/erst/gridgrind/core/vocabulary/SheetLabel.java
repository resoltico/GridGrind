package dev.erst.gridgrind.core.vocabulary;

import java.util.Objects;

/** Observed worksheet text, including labels that violate Excel's authoring-name rules. */
public record SheetLabel(String value) implements WireString {
  public SheetLabel {
    Objects.requireNonNull(value, "sheetName must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException("sheetName must not be blank");
    }
  }

  @Override
  public String toString() {
    return value;
  }
}
