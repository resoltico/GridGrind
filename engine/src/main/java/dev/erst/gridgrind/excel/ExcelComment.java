package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Immutable workbook-core plain-text comment used for cell comment authoring and analysis. */
public record ExcelComment(String text, String author, boolean visible) {
  public ExcelComment {
    Objects.requireNonNull(text, "text must not be null");
    Objects.requireNonNull(author, "author must not be null");
    if (text.isBlank()) {
      throw new IllegalArgumentException("text must not be blank");
    }
    if (author.isBlank()) {
      throw new IllegalArgumentException("author must not be blank");
    }
  }
}
