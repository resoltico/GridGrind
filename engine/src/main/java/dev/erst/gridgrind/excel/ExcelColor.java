package dev.erst.gridgrind.excel;

import java.util.Objects;
import java.util.Optional;

/** Mutable-workbook color payload preserving RGB, theme, indexed, and tint semantics. */
public sealed interface ExcelColor permits ExcelColor.Rgb, ExcelColor.Theme, ExcelColor.Indexed {
  /** Optional tint adjustment applied to the base color reference. */
  Optional<Double> tint();

  /** Creates one mutable-workbook color carrying only explicit RGB data. */
  static Rgb rgb(String rgb) {
    return new Rgb(rgb, Optional.empty());
  }

  /** Creates one RGB-backed mutable color plus tint metadata. */
  static Rgb rgb(String rgb, Optional<Double> tint) {
    return new Rgb(rgb, tint);
  }

  /** Creates one RGB-backed mutable color plus tint metadata. */
  static Rgb rgb(String rgb, double tint) {
    return new Rgb(rgb, Optional.of(tint));
  }

  /** Creates one theme-backed mutable color with no tint adjustment. */
  static Theme theme(int theme) {
    return new Theme(theme, Optional.empty());
  }

  /** Creates one theme-backed mutable color plus tint metadata. */
  static Theme theme(int theme, Optional<Double> tint) {
    return new Theme(theme, tint);
  }

  /** Creates one theme-backed mutable color plus tint metadata. */
  static Theme theme(int theme, double tint) {
    return new Theme(theme, Optional.of(tint));
  }

  /** Creates one indexed-palette mutable color with no tint adjustment. */
  static Indexed indexed(int indexed) {
    return new Indexed(indexed, Optional.empty());
  }

  /** Creates one indexed-palette mutable color plus tint metadata. */
  static Indexed indexed(int indexed, Optional<Double> tint) {
    return new Indexed(indexed, tint);
  }

  /** Creates one indexed-palette mutable color plus tint metadata. */
  static Indexed indexed(int indexed, double tint) {
    return new Indexed(indexed, Optional.of(tint));
  }

  /** RGB-backed mutable workbook color. */
  record Rgb(String rgb, Optional<Double> tint) implements ExcelColor {
    public Rgb {
      rgb = ExcelRgbColorSupport.requireRgbHex(rgb, "rgb");
      tint = requireFinite(tint, "tint");
    }
  }

  /** Theme-backed mutable workbook color. */
  record Theme(Integer theme, Optional<Double> tint) implements ExcelColor {
    public Theme {
      requireNonNegative(theme, "theme");
      tint = requireFinite(tint, "tint");
    }
  }

  /** Indexed-palette mutable workbook color. */
  record Indexed(Integer indexed, Optional<Double> tint) implements ExcelColor {
    public Indexed {
      requireNonNegative(indexed, "indexed");
      tint = requireFinite(tint, "tint");
    }
  }

  private static void requireNonNegative(Integer value, String fieldName) {
    int required = Objects.requireNonNull(value, fieldName + " must not be null");
    if (required < 0) {
      throw new IllegalArgumentException(fieldName + " must not be negative");
    }
  }

  private static Optional<Double> requireFinite(Optional<Double> value, String fieldName) {
    Optional<Double> required = Objects.requireNonNull(value, fieldName + " must not be null");
    required.ifPresent(
        finite -> {
          if (!Double.isFinite(finite)) {
            throw new IllegalArgumentException(fieldName + " must be finite");
          }
        });
    return required;
  }
}
