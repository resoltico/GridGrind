package dev.erst.gridgrind.core.vocabulary;

import java.util.Objects;

/** A factual stored range label, including malformed references that reports must preserve. */
public record RangeLabel(String value) implements WireString {
  public RangeLabel {
    Objects.requireNonNull(value, "range must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException("range must not be blank");
    }
  }

  @Override
  public String toString() {
    return value;
  }
}
