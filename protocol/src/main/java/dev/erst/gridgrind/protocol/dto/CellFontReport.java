package dev.erst.gridgrind.protocol.dto;

import java.util.Objects;

/** Effective font facts reported with every analyzed cell. */
public record CellFontReport(
    boolean bold,
    boolean italic,
    String fontName,
    FontHeightReport fontHeight,
    String fontColor,
    boolean underline,
    boolean strikeout) {
  public CellFontReport {
    Objects.requireNonNull(fontName, "fontName must not be null");
    Objects.requireNonNull(fontHeight, "fontHeight must not be null");
    fontColor = ProtocolRgbColorSupport.normalizeRgbHex(fontColor, "fontColor");
  }
}
