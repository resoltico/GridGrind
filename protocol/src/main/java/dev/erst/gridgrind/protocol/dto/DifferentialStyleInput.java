package dev.erst.gridgrind.protocol.dto;

import java.util.Objects;
import java.util.stream.Stream;

/** Protocol-facing differential style payload for authored conditional-formatting rules. */
public record DifferentialStyleInput(
    String numberFormat,
    Boolean bold,
    Boolean italic,
    FontHeightInput fontHeight,
    String fontColor,
    Boolean underline,
    Boolean strikeout,
    String fillColor,
    DifferentialBorderInput border) {
  public DifferentialStyleInput {
    if (numberFormat != null && numberFormat.isBlank()) {
      throw new IllegalArgumentException("numberFormat must not be blank");
    }
    fontColor = ProtocolRgbColorSupport.normalizeRgbHex(fontColor, "fontColor");
    fillColor = ProtocolRgbColorSupport.normalizeRgbHex(fillColor, "fillColor");
    if (hasNoStyleAttributes(
        numberFormat,
        bold,
        italic,
        fontHeight,
        fontColor,
        underline,
        strikeout,
        fillColor,
        border)) {
      throw new IllegalArgumentException("style must set at least one attribute");
    }
  }

  private static boolean hasNoStyleAttributes(Object... attributes) {
    return Stream.of(attributes).allMatch(Objects::isNull);
  }
}
