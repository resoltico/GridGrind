package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Immutable workbook-core comment used for cell comment authoring and analysis. */
public record ExcelComment(
    String text, String author, boolean visible, ExcelRichText runs, ExcelCommentAnchor anchor) {
  /** Creates a plain-text comment without rich-text runs or an explicit anchor override. */
  public ExcelComment(String text, String author, boolean visible) {
    this(text, author, visible, null, null);
  }

  public ExcelComment {
    Objects.requireNonNull(text, "text must not be null");
    Objects.requireNonNull(author, "author must not be null");
    if (text.isBlank()) {
      throw new IllegalArgumentException("text must not be blank");
    }
    if (author.isBlank()) {
      throw new IllegalArgumentException("author must not be blank");
    }
    if (runs != null && !text.equals(runs.plainText())) {
      throw new IllegalArgumentException("comment run text must concatenate to the plain text");
    }
  }
}
