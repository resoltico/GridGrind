package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.style.CellColorReport;
import java.util.Objects;
import org.jspecify.annotations.Nullable;

/** Immutable snapshot of the font currently applied to a cell. */
public record ExcelCellFontSnapshot(
    boolean bold,
    boolean italic,
    String fontName,
    ExcelFontHeight fontHeight,
    @Nullable CellColorReport fontColor,
    boolean underline,
    boolean strikeout) {
  public ExcelCellFontSnapshot {
    Objects.requireNonNull(fontName, "fontName must not be null");
    Objects.requireNonNull(fontHeight, "fontHeight must not be null");
  }
}
