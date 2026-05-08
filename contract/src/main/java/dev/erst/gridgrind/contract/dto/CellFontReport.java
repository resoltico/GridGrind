package dev.erst.gridgrind.contract.dto;

import java.util.Objects;
import org.jspecify.annotations.Nullable;

/** Effective font facts reported with every analyzed cell. */
public record CellFontReport(
    boolean bold,
    boolean italic,
    String fontName,
    FontHeightReport fontHeight,
    @Nullable CellColorReport fontColor,
    boolean underline,
    boolean strikeout) {
  public CellFontReport {
    Objects.requireNonNull(fontName, "fontName must not be null");
    Objects.requireNonNull(fontHeight, "fontHeight must not be null");
  }
}
