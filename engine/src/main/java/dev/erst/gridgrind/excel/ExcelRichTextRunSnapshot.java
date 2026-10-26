package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.core.model.style.CellFontReport;
import java.util.Objects;

/** Immutable factual snapshot of one rich-text run stored in a cell. */
public record ExcelRichTextRunSnapshot(String text, CellFontReport font) {
  public ExcelRichTextRunSnapshot {
    Objects.requireNonNull(text, "text must not be null");
    Objects.requireNonNull(font, "font must not be null");
    if (text.isEmpty()) {
      throw new IllegalArgumentException("text must not be empty");
    }
  }
}
