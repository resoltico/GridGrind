package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import java.util.Objects;
import java.util.Optional;

/** Protocol-facing color payload preserving RGB, theme, indexed, and tint semantics. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "kind")
@JsonSubTypes({
  @JsonSubTypes.Type(value = ColorInput.Rgb.class, name = "RGB"),
  @JsonSubTypes.Type(value = ColorInput.Theme.class, name = "THEME"),
  @JsonSubTypes.Type(value = ColorInput.Indexed.class, name = "INDEXED")
})
public sealed interface ColorInput permits ColorInput.Rgb, ColorInput.Theme, ColorInput.Indexed {
  /** Optional tint adjustment applied to the base color reference. */
  Optional<Double> tint();

  /** Creates one protocol color carrying only explicit RGB data. */
  static Rgb rgb(String rgb) {
    return new Rgb(rgb, Optional.empty());
  }

  /** Creates one protocol color carrying explicit RGB data plus tint metadata. */
  static Rgb rgb(String rgb, double tint) {
    return new Rgb(rgb, Optional.of(tint));
  }

  /** Creates one protocol color referencing one workbook theme slot. */
  static Theme theme(int theme) {
    return new Theme(theme, Optional.empty());
  }

  /** Creates one protocol color referencing one workbook theme slot plus tint metadata. */
  static Theme theme(int theme, double tint) {
    return new Theme(theme, Optional.of(tint));
  }

  /** Creates one protocol color referencing one indexed workbook palette slot. */
  static Indexed indexed(int indexed) {
    return new Indexed(indexed, Optional.empty());
  }

  /** Creates one protocol color referencing one indexed workbook palette slot plus tint. */
  static Indexed indexed(int indexed, double tint) {
    return new Indexed(indexed, Optional.of(tint));
  }

  /** Explicit RGB color reference. */
  record Rgb(String rgb, @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<Double> tint)
      implements ColorInput {
    public Rgb {
      rgb = ProtocolRgbColorSupport.requireRgbHex(rgb, "rgb");
      tint = requireFinite(tint, "tint");
    }
  }

  /** Theme-slot color reference. */
  record Theme(int theme, @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<Double> tint)
      implements ColorInput {
    public Theme {
      requireNonNegative(theme, "theme");
      tint = requireFinite(tint, "tint");
    }
  }

  /** Indexed-palette color reference. */
  record Indexed(int indexed, @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<Double> tint)
      implements ColorInput {
    public Indexed {
      requireNonNegative(indexed, "indexed");
      tint = requireFinite(tint, "tint");
    }
  }

  private static void requireNonNegative(int value, String fieldName) {
    if (value < 0) {
      throw new IllegalArgumentException(fieldName + " must not be negative");
    }
  }

  private static Optional<Double> requireFinite(Optional<Double> value, String fieldName) {
    Optional<Double> normalized = Objects.requireNonNull(value, fieldName + " must not be null");
    if (normalized.isPresent() && !Double.isFinite(normalized.orElseThrow())) {
      throw new IllegalArgumentException(fieldName + " must be finite");
    }
    return normalized;
  }
}
