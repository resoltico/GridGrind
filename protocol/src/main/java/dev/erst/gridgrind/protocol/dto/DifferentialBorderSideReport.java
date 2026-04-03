package dev.erst.gridgrind.protocol.dto;

import java.util.Objects;

/** Protocol-facing factual report for one conditional-formatting differential border side. */
public record DifferentialBorderSideReport(BorderStyle style, String color) {
  public DifferentialBorderSideReport {
    Objects.requireNonNull(style, "style must not be null");
    color = ProtocolRgbColorSupport.normalizeRgbHex(color, "color");
  }
}
