package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Mutable-workbook color payload preserving RGB, theme, indexed, and tint semantics. */
public sealed interface ExcelColor permits ExcelColor.Rgb, ExcelColor.Theme, ExcelColor.Indexed {
  /** Optional tint adjustment applied to the base color reference. */
  Double tint();

  /** Creates one mutable-workbook color carrying only explicit RGB data. */
  static Rgb rgb(String rgb) {
    return new Rgb(rgb, null);
  }

  /** Creates one RGB-backed mutable color plus tint metadata. */
  static Rgb rgb(String rgb, Double tint) {
    return new Rgb(rgb, tint);
  }

  /** Creates one theme-backed mutable color with no tint adjustment. */
  static Theme theme(int theme) {
    return new Theme(theme, null);
  }

  /** Creates one theme-backed mutable color plus tint metadata. */
  static Theme theme(int theme, Double tint) {
    return new Theme(theme, tint);
  }

  /** Creates one indexed-palette mutable color with no tint adjustment. */
  static Indexed indexed(int indexed) {
    return new Indexed(indexed, null);
  }

  /** Creates one indexed-palette mutable color plus tint metadata. */
  static Indexed indexed(int indexed, Double tint) {
    return new Indexed(indexed, tint);
  }

  /** RGB-backed mutable workbook color. */
  record Rgb(String rgb, Double tint) implements ExcelColor {
    public Rgb {
      rgb = ExcelRgbColorSupport.requireRgbHex(rgb, "rgb");
      requireFiniteOrNull(tint, "tint");
    }
  }

  /** Theme-backed mutable workbook color. */
  record Theme(Integer theme, Double tint) implements ExcelColor {
    public Theme {
      requireNonNegative(theme, "theme");
      requireFiniteOrNull(tint, "tint");
    }
  }

  /** Indexed-palette mutable workbook color. */
  record Indexed(Integer indexed, Double tint) implements ExcelColor {
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
