package dev.erst.gridgrind.core.vocabulary;

import java.util.Locale;
import java.util.Objects;

/** An RGB literal in canonical uppercase #RRGGBB form. */
public record RgbHex(String value) implements WireString {
  public RgbHex {
    Objects.requireNonNull(value, "rgb must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException("rgb must not be blank");
    }
    if (!value.matches(ProtocolConstraintValues.RGB_HEX_PATTERN)) {
      throw new IllegalArgumentException("rgb must match #RRGGBB");
    }
    value = value.toUpperCase(Locale.ROOT);
  }

  @Override
  public String toString() {
    return value;
  }
}
