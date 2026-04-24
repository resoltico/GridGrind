package dev.erst.gridgrind.excel;

import java.util.Locale;
import java.util.Objects;
import java.util.Optional;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/** Shared RGB color normalization and POI conversion helpers for workbook style subsystems. */
final class ExcelRgbColorSupport {
  private ExcelRgbColorSupport() {}

  /** Normalizes one optional {@code #RRGGBB} color literal for storage and comparison. */
  static Optional<String> normalizeRgbHex(String color, String fieldName) {
    if (color == null) {
      return Optional.empty();
    }
    Objects.requireNonNull(fieldName, "fieldName must not be null");
    if (color.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    if (!color.matches("^#[0-9A-Fa-f]{6}$")) {
      throw new IllegalArgumentException(fieldName + " must match #RRGGBB");
    }
    return Optional.of(color.toUpperCase(Locale.ROOT));
  }

  /** Converts one POI workbook color into {@code #RRGGBB}, or null when no RGB is available. */
  static Optional<String> toRgbHex(XSSFColor color) {
    if (color == null) {
      return Optional.empty();
    }
    byte[] rgb = color.getRGB();
    if (rgb == null || rgb.length != 3) {
      return Optional.empty();
    }
    return Optional.of("#%02X%02X%02X".formatted(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF));
  }

  /** Converts one normalized {@code #RRGGBB} literal into the matching POI workbook color. */
  static XSSFColor toXssfColor(XSSFWorkbook workbook, String color) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    String normalized =
        normalizeRgbHex(color, "color")
            .orElseThrow(() -> new IllegalArgumentException("color must not be null"));
    return new XSSFColor(
        new byte[] {
          (byte) Integer.parseInt(normalized.substring(1, 3), 16),
          (byte) Integer.parseInt(normalized.substring(3, 5), 16),
          (byte) Integer.parseInt(normalized.substring(5, 7), 16)
        },
        workbook.getStylesSource().getIndexedColors());
  }
}
