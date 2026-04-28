package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Immutable factual workbook color preserving RGB, theme, indexed, and tint semantics. */
public sealed interface ExcelColorSnapshot
    permits ExcelColorSnapshot.Rgb, ExcelColorSnapshot.Theme, ExcelColorSnapshot.Indexed {
  /** Optional tint adjustment applied to the base color reference. */
  Double tint();

  /** Returns one snapshot carrying only explicit RGB data. */
  static Rgb rgb(String rgb) {
    return new Rgb(rgb, null);
  }

  /** Returns one RGB-backed snapshot plus tint metadata. */
  static Rgb rgb(String rgb, Double tint) {
    return new Rgb(rgb, tint);
  }

  /** Returns one theme-backed snapshot with no tint adjustment. */
  static Theme theme(int theme) {
    return new Theme(theme, null);
  }

  /** Returns one theme-backed snapshot plus tint metadata. */
  static Theme theme(int theme, Double tint) {
    return new Theme(theme, tint);
  }

  /** Returns one indexed-palette snapshot with no tint adjustment. */
  static Indexed indexed(int indexed) {
    return new Indexed(indexed, null);
  }

  /** Returns one indexed-palette snapshot plus tint metadata. */
  static Indexed indexed(int indexed, Double tint) {
    return new Indexed(indexed, tint);
  }

  /** RGB-backed workbook color snapshot. */
  record Rgb(String rgb, Double tint) implements ExcelColorSnapshot {
    public Rgb {
      rgb = ExcelRgbColorSupport.requireRgbHex(rgb, "rgb");
      requireFiniteOrNull(tint, "tint");
    }
  }

  /** Theme-backed workbook color snapshot. */
  record Theme(Integer theme, Double tint) implements ExcelColorSnapshot {
    public Theme {
      requireNonNegative(theme, "theme");
      requireFiniteOrNull(tint, "tint");
    }
  }

  /** Indexed-palette workbook color snapshot. */
  record Indexed(Integer indexed, Double tint) implements ExcelColorSnapshot {
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
