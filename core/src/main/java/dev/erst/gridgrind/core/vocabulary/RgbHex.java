package dev.erst.gridgrind.core.vocabulary;

import java.util.Locale;

/** An RGB literal in canonical uppercase #RRGGBB form. */
public record RgbHex(String value) implements WireString {
  public RgbHex {
    value = TextValidation.requireNonBlank(value, "rgb");
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
