package dev.erst.gridgrind.core.vocabulary;

import java.util.Objects;

/** A stored table label, including empty or malformed names retained by factual readback. */
public record TableLabel(String value) implements WireString {
  public TableLabel {
    Objects.requireNonNull(value, "name must not be null");
  }

  @Override
  public String toString() {
    return value;
  }
}
