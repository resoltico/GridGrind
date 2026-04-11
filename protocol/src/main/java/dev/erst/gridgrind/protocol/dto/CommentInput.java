package dev.erst.gridgrind.protocol.dto;

import java.util.Objects;

/** Protocol-facing plain-text comment payload used by cell comment operations. */
public record CommentInput(
    String text,
    String author,
    Boolean visible,
    java.util.List<RichTextRunInput> runs,
    CommentAnchorInput anchor) {
  /** Creates a plain-text comment payload without rich-text runs or anchor overrides. */
  public CommentInput(String text, String author, Boolean visible) {
    this(text, author, visible, null, null);
  }

  public CommentInput {
    text = requireNonBlank(text, "text");
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
      String joined =
          runs.stream().map(RichTextRunInput::text).collect(java.util.stream.Collectors.joining());
      if (!text.equals(joined)) {
        throw new IllegalArgumentException("comment runs must concatenate to the plain text");
      }
    }
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
