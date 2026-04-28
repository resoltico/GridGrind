package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.Objects;

/** Protocol-facing color payload preserving RGB, theme, indexed, and tint semantics. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "kind")
@JsonSubTypes({
  @JsonSubTypes.Type(value = ColorInput.Rgb.class, name = "RGB"),
  @JsonSubTypes.Type(value = ColorInput.Theme.class, name = "THEME"),
  @JsonSubTypes.Type(value = ColorInput.Indexed.class, name = "INDEXED")
})
public sealed interface ColorInput permits ColorInput.Rgb, ColorInput.Theme, ColorInput.Indexed {
  /** Optional tint adjustment applied to the base color reference. */
  Double tint();

  /** Creates one protocol color carrying only explicit RGB data. */
  static Rgb rgb(String rgb) {
    return new Rgb(rgb, null);
  }

  /** Creates one protocol color carrying explicit RGB data plus tint metadata. */
  static Rgb rgb(String rgb, Double tint) {
    return new Rgb(rgb, tint);
  }

  /** Creates one protocol color referencing one workbook theme slot. */
  static Theme theme(int theme) {
    return new Theme(theme, null);
  }

  /** Creates one protocol color referencing one workbook theme slot plus tint metadata. */
  static Theme theme(int theme, Double tint) {
    return new Theme(theme, tint);
  }

  /** Creates one protocol color referencing one indexed workbook palette slot. */
  static Indexed indexed(int indexed) {
    return new Indexed(indexed, null);
  }

  /** Creates one protocol color referencing one indexed workbook palette slot plus tint. */
  static Indexed indexed(int indexed, Double tint) {
    return new Indexed(indexed, tint);
  }

  /** Explicit RGB color reference. */
  record Rgb(String rgb, Double tint) implements ColorInput {
    public Rgb {
      rgb = ProtocolRgbColorSupport.requireRgbHex(rgb, "rgb");
      requireFiniteOrNull(tint, "tint");
    }
  }

  /** Theme-slot color reference. */
  record Theme(Integer theme, Double tint) implements ColorInput {
    public Theme {
      requireNonNegative(theme, "theme");
      requireFiniteOrNull(tint, "tint");
    }
  }

  /** Indexed-palette color reference. */
  record Indexed(Integer indexed, Double tint) implements ColorInput {
    public Indexed {
      requireNonNegative(indexed, "indexed");
      requireFiniteOrNull(tint, "tint");
    }
  }

  private static void requireNonNegative(Integer value, String fieldName) {
    int required = Objects.requireNonNull(value, fieldName + " must not be null");
    if (required < 0) {
      throw new IllegalArgumentException(fieldName + " must not be negative");
    }
  }

  private static void requireFiniteOrNull(Double value, String fieldName) {
    if (value != null && !Double.isFinite(value)) {
      throw new IllegalArgumentException(fieldName + " must be finite");
    }
  }
}
