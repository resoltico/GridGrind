package dev.erst.gridgrind.core.model.style;

import dev.erst.gridgrind.core.metadata.CatalogType;
import java.util.Objects;
import java.util.Optional;

/** Effective font facts reported with every analyzed cell. */
@CatalogType(id = "CellFontReport", summary = "Exact cell-font report.", displayOrder = 80)
public record CellFontReport(
    boolean bold,
    boolean italic,
    String fontName,
    FontHeightReport fontHeight,
    Optional<CellColorReport> fontColor,
    boolean underline,
    boolean strikeout) {
  public CellFontReport {
    Objects.requireNonNull(fontName, "fontName must not be null");
    Objects.requireNonNull(fontHeight, "fontHeight must not be null");
    Objects.requireNonNull(fontColor, "fontColor must not be null");
  }
}
