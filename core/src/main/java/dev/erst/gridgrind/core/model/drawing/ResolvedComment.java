package dev.erst.gridgrind.core.model.drawing;

import dev.erst.gridgrind.core.model.style.ResolvedRichText;
import java.util.Objects;
import java.util.Optional;

/** Materialized comment text and rich runs ready for workbook authoring. */
public record ResolvedComment(
    String text,
    String author,
    boolean visible,
    Optional<ResolvedRichText> runs,
    Optional<CommentAnchorSpec> anchor) {
  /** Creates a plain-text comment without rich-text runs or an explicit anchor override. */
  public ResolvedComment(String text, String author, boolean visible) {
    this(text, author, visible, Optional.empty(), Optional.empty());
  }

  public ResolvedComment {
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
