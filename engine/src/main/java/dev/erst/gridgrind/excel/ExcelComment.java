package dev.erst.gridgrind.excel;

import java.util.Objects;
import java.util.Optional;

/** Immutable workbook-core comment used for cell comment authoring and analysis. */
public record ExcelComment(
    String text,
    String author,
    boolean visible,
    Optional<ExcelRichText> runs,
    Optional<ExcelCommentAnchor> anchor) {
  /** Creates a plain-text comment without rich-text runs or an explicit anchor override. */
  public ExcelComment(String text, String author, boolean visible) {
    this(text, author, visible, Optional.empty(), Optional.empty());
  }

  public ExcelComment {
    Objects.requireNonNull(text, "text must not be null");
    Objects.requireNonNull(author, "author must not be null");
    Objects.requireNonNull(runs, "runs must not be null");
    Objects.requireNonNull(anchor, "anchor must not be null");
    if (text.isBlank()) {
      throw new IllegalArgumentException("text must not be blank");
    }
    if (author.isBlank()) {
      throw new IllegalArgumentException("author must not be blank");
    }
    if (runs.isPresent() && !text.equals(runs.orElseThrow().plainText())) {
      throw new IllegalArgumentException("comment run text must concatenate to the plain text");
    }
  }
}
