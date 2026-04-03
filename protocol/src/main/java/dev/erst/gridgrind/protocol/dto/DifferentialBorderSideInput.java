package dev.erst.gridgrind.protocol.dto;

import java.util.Objects;

/** Protocol-facing definition for one conditional-formatting differential border side. */
public record DifferentialBorderSideInput(BorderStyle style, String color) {
  public DifferentialBorderSideInput {
    Objects.requireNonNull(style, "style must not be null");
    color = ProtocolRgbColorSupport.normalizeRgbHex(color, "color");
  }
}
