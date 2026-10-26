package dev.erst.gridgrind.core.model.style;

import com.fasterxml.jackson.annotation.JsonInclude;
import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import java.util.Objects;
import java.util.Optional;

/** Protocol-facing color payload preserving RGB, theme, indexed, and tint semantics. */
@Union(discriminator = "type", catalog = true, displayOrder = 55)
public sealed interface ColorInput
    permits dev.erst.gridgrind.core.model.style.ColorInput.Rgb,
        dev.erst.gridgrind.core.model.style.ColorInput.Theme,
        dev.erst.gridgrind.core.model.style.ColorInput.Indexed {
  /** Optional tint adjustment applied to the base color reference. */
  Optional<Double> tint();

  /** Creates one protocol color carrying only explicit RGB data. */
  static Rgb rgb(RgbHex rgb) {
    return new Rgb(rgb, Optional.empty());
  }

  /** Creates one protocol color carrying explicit RGB data plus tint metadata. */
  static Rgb rgb(RgbHex rgb, double tint) {
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
  @Variant(
      id = "RGB",
      summary = "Write one explicit RGB color reference plus optional tint metadata.")
  record Rgb(
      @dev.erst.gridgrind.core.metadata.FieldRule(
              kind = dev.erst.gridgrind.core.metadata.FieldRule.Kind.STRING_PATTERN,
              text = dev.erst.gridgrind.core.vocabulary.ProtocolConstraintValues.RGB_HEX_PATTERN)
          RgbHex rgb,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<Double> tint)
      implements ColorInput {
    public Rgb {
      java.util.Objects.requireNonNull(rgb, "rgb must not be null");
      tint = requireFinite(tint, "tint");
    }
  }

  /** Theme-slot color reference. */
  @Variant(
      id = "THEME",
      summary = "Write one workbook theme-slot color reference plus optional tint metadata.")
  record Theme(int theme, @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<Double> tint)
      implements ColorInput {
    public Theme {
      requireNonNegative(theme, "theme");
      tint = requireFinite(tint, "tint");
    }
  }

  /** Indexed-palette color reference. */
  @Variant(
      id = "INDEXED",
      summary = "Write one indexed-palette color reference plus optional tint metadata.")
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
