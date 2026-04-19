package dev.erst.gridgrind.contract.dto;

import java.util.Locale;

/** Factual workbook color preserving RGB, theme, indexed, and tint semantics. */
public record CellColorReport(String rgb, Integer theme, Integer indexed, Double tint) {
  public CellColorReport {
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
      throw new IllegalArgumentException(
          "color report must expose rgb, theme, or indexed semantics");
    }
  }

  /** Returns one color report carrying only explicit RGB data. */
  public CellColorReport(String rgb) {
    this(rgb, null, null, null);
  }

  private static String normalizeRgbHex(String rgb) {
    if (rgb == null) {
      return null;
    }
    if (rgb.isBlank()) {
      throw new IllegalArgumentException("rgb must not be blank");
    }
    if (!rgb.matches("^#[0-9A-Fa-f]{6}$")) {
      throw new IllegalArgumentException("rgb must match #RRGGBB");
    }
    return rgb.toUpperCase(Locale.ROOT);
  }
}
