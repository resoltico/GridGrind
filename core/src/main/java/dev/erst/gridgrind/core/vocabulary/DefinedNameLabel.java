package dev.erst.gridgrind.core.vocabulary;

import java.util.Objects;

/** A nonblank defined-name label observed in a workbook, without authoring restrictions. */
public record DefinedNameLabel(String value) implements WireString {
  public DefinedNameLabel {
    Objects.requireNonNull(value, "name must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException("name must not be blank");
    }
  }

  @Override
  public String toString() {
    return value;
  }
}
