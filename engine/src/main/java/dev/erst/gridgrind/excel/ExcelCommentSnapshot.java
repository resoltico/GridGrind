package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Immutable factual comment metadata preserving plain text, runs, visibility, and anchor data. */
public record ExcelCommentSnapshot(
    String text,
    String author,
    boolean visible,
    ExcelRichTextSnapshot runs,
    ExcelCommentAnchorSnapshot anchor) {
  public ExcelCommentSnapshot {
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

  /** Returns the plain-text authoring view of this factual comment snapshot. */
  public ExcelComment toPlainComment() {
    return new ExcelComment(text, author, visible);
  }
}
