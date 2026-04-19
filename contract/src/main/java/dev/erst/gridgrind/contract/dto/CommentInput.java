package dev.erst.gridgrind.contract.dto;

import dev.erst.gridgrind.contract.source.TextSourceInput;
import java.util.Objects;

/** Protocol-facing plain-text comment payload used by cell comment operations. */
public record CommentInput(
    TextSourceInput text,
    String author,
    Boolean visible,
    java.util.List<RichTextRunInput> runs,
    CommentAnchorInput anchor) {
  /** Creates a plain-text comment payload without rich-text runs or anchor overrides. */
  public CommentInput(TextSourceInput text, String author, Boolean visible) {
    this(text, author, visible, null, null);
  }

  public CommentInput {
    text = requireNonBlankSource(text, "text");
    author = requireNonBlank(author, "author");
    visible = visible == null ? Boolean.FALSE : visible;
    if (runs != null) {
      runs = java.util.List.copyOf(runs);
      if (runs.isEmpty()) {
        throw new IllegalArgumentException("runs must not be empty when provided");
      }
      for (RichTextRunInput run : runs) {
        Objects.requireNonNull(run, "runs must not contain null values");
      }
      validateRunsAgainstInlineText(text, runs);
    }
  }

  private static void validateRunsAgainstInlineText(
      TextSourceInput text, java.util.List<RichTextRunInput> runs) {
    if (!(text instanceof TextSourceInput.Inline inlineText)) {
      return;
    }
    if (runs.stream().anyMatch(run -> !(run.source() instanceof TextSourceInput.Inline))) {
      return;
    }
    String joined =
        runs.stream()
            .map(RichTextRunInput::source)
            .map(TextSourceInput.Inline.class::cast)
            .map(TextSourceInput.Inline::text)
            .collect(java.util.stream.Collectors.joining());
    if (!inlineText.text().equals(joined)) {
      throw new IllegalArgumentException("comment runs must concatenate to the plain text");
    }
  }

  private static TextSourceInput requireNonBlankSource(TextSourceInput value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value instanceof TextSourceInput.Inline inline && inline.text().isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
