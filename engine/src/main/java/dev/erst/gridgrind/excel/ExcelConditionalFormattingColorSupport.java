package dev.erst.gridgrind.excel;

import java.util.Objects;
import java.util.Optional;
import org.jspecify.annotations.Nullable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTColor;

/** Encodes and decodes differential-style RGB payloads carried through OOXML color nodes. */
final class ExcelConditionalFormattingColorSupport {
  private ExcelConditionalFormattingColorSupport() {}

  static byte[] argbBytes(String rgbHex) {
    String normalized =
        ExcelRgbColorSupport.normalizeRgbHex(rgbHex, "rgbHex")
            .orElseThrow(() -> new IllegalArgumentException("rgbHex must not be null"));
    return new byte[] {
      (byte) 0xFF,
      (byte) Integer.parseInt(normalized.substring(1, 3), 16),
      (byte) Integer.parseInt(normalized.substring(3, 5), 16),
      (byte) Integer.parseInt(normalized.substring(5, 7), 16)
    };
  }

  static void setColor(CTColor color, String rgbHex) {
    Objects.requireNonNull(color, "color must not be null");
    color.setRgb(argbBytes(rgbHex));
  }

  static @Nullable String rgbHexFromCtColor(@Nullable CTColor color) {
    return optionalRgbHexFromCtColor(color).orElse(null);
  }

  static Optional<String> optionalRgbHexFromCtColor(@Nullable CTColor color) {
    if (color == null || !color.isSetRgb()) {
      return Optional.empty();
    }
    byte[] rgb = color.getRgb();
    if (rgb.length == 4) {
      return Optional.of("#%02X%02X%02X".formatted(rgb[1] & 0xFF, rgb[2] & 0xFF, rgb[3] & 0xFF));
    }
    if (rgb.length == 3) {
      return Optional.of("#%02X%02X%02X".formatted(rgb[0] & 0xFF, rgb[1] & 0xFF, rgb[2] & 0xFF));
    }
    return Optional.empty();
  }
}
