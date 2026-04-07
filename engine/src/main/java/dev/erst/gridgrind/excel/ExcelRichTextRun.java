package dev.erst.gridgrind.excel;

import java.util.Objects;

/** One ordered rich-text run authored for a string cell value. */
public record ExcelRichTextRun(String text, ExcelCellFont font) {
  public ExcelRichTextRun {
    Objects.requireNonNull(text, "text must not be null");
    if (text.isEmpty()) {
      throw new IllegalArgumentException("text must not be empty");
    }
  }
}
