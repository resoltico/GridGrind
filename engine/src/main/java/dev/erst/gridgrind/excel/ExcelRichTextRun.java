package dev.erst.gridgrind.excel;

import java.util.Objects;
import java.util.Optional;

/** One ordered rich-text run authored for a string cell value. */
public record ExcelRichTextRun(String text, Optional<ExcelCellFont> font) {
  public ExcelRichTextRun {
    Objects.requireNonNull(text, "text must not be null");
    Objects.requireNonNull(font, "font must not be null");
    if (text.isEmpty()) {
      throw new IllegalArgumentException("text must not be empty");
    }
  }
}
