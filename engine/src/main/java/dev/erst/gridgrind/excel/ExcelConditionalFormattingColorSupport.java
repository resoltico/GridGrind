package dev.erst.gridgrind.excel;

import java.util.Objects;
import java.util.Optional;
import org.jspecify.annotations.Nullable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor;

/** Encodes and decodes differential-style colors without losing their OOXML reference semantics. */
final class ExcelConditionalFormattingColorSupport {
  private ExcelConditionalFormattingColorSupport() {}

  /**
   * Writes one owned workbook color without collapsing theme, indexed, or tint semantics to RGB.
   */
  static void setColor(CTColor target, ExcelColor color) {
    Objects.requireNonNull(target, "target must not be null");
    Objects.requireNonNull(color, "color must not be null");
    switch (color) {
      case ExcelColor.Rgb rgb -> {
        target.setRgb(argbBytes(rgb));
        rgb.tint().ifPresent(target::setTint);
      }
      case ExcelColor.Theme theme -> {
        target.setTheme(theme.theme());
        theme.tint().ifPresent(target::setTint);
      }
      case ExcelColor.Indexed indexed -> {
        target.setIndexed(indexed.indexed());
        indexed.tint().ifPresent(target::setTint);
      }
    }
  }

  /** Reads one OOXML color into the owned color vocabulary, preserving tint and reference kind. */
  static Optional<ExcelColor> optionalColor(@Nullable CTColor color) {
    if (color == null) {
      return Optional.empty();
    }
    Optional<Double> tint = color.isSetTint() ? Optional.of(color.getTint()) : Optional.empty();
    if (color.isSetRgb()) {
      return rgbHex(color.getRgb()).map(value -> ExcelColor.rgb(value, tint));
    }
    if (color.isSetTheme()) {
      return Optional.of(ExcelColor.theme(Math.toIntExact(color.getTheme()), tint));
    }
    if (color.isSetIndexed()) {
      return Optional.of(ExcelColor.indexed(Math.toIntExact(color.getIndexed()), tint));
    }
    return Optional.empty();
  }

  private static byte[] argbBytes(ExcelColor.Rgb color) {
    String rgb = color.rgb();
    return new byte[] {
      (byte) 0xFF,
      (byte) Integer.parseInt(rgb.substring(1, 3), 16),
      (byte) Integer.parseInt(rgb.substring(3, 5), 16),
      (byte) Integer.parseInt(rgb.substring(5, 7), 16)
    };
  }

  private static Optional<String> rgbHex(byte[] rgb) {
    if (rgb.length == 4) {
      return Optional.of("#%02X%02X%02X".formatted(rgb[1] & 0xFF, rgb[2] & 0xFF, rgb[3] & 0xFF));
    }
    if (rgb.length == 3) {
      return Optional.of("#%02X%02X%02X".formatted(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF));
    }
    return Optional.empty();
  }
}
