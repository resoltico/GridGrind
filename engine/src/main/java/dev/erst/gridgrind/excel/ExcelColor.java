package dev.erst.gridgrind.excel;

import java.util.Locale;
import java.util.Objects;

/** Mutable-workbook color payload preserving RGB, theme, indexed, and tint semantics. */
public record ExcelColor(String rgb, Integer theme, Integer indexed, Double tint) {
  /** Creates one mutable-workbook color carrying only explicit RGB data. */
  public ExcelColor(String rgb) {
    this(rgb, null, null, null);
  }

  public ExcelColor {
    rgb = normalizeRgbHex(rgb);
    if (theme != null && theme < 0) {
      throw new IllegalArgumentException("theme must not be negative");
    }
    if (indexed != null && indexed < 0) {
      throw new IllegalArgumentException("indexed must not be negative");
    }
    if (tint != null && !Double.isFinite(tint)) {
      throw new IllegalArgumentException("tint must be finite");
    }
    if (rgb == null && theme == null && indexed == null) {
      throw new IllegalArgumentException("color must expose rgb, theme, or indexed semantics");
    }
  }

  private static String normalizeRgbHex(String rgb) {
    if (rgb == null) {
      return null;
    }
    Objects.requireNonNull(rgb, "rgb must not be null");
    if (rgb.isBlank()) {
      throw new IllegalArgumentException("rgb must not be blank");
    }
    if (!rgb.matches("^#[0-9A-Fa-f]{6}$")) {
      throw new IllegalArgumentException("rgb must match #RRGGBB");
    }
    return rgb.toUpperCase(Locale.ROOT);
  }
}
