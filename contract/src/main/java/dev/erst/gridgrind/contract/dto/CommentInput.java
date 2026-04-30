package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import java.util.Objects;
import java.util.Optional;

/** Protocol-facing plain-text comment payload used by cell comment operations. */
public record CommentInput(
    TextSourceInput text,
    String author,
    boolean visible,
    @com.fasterxml.jackson.annotation.JsonInclude(
            com.fasterxml.jackson.annotation.JsonInclude.Include.NON_ABSENT)
        Optional<java.util.List<RichTextRunInput>> runs,
    @com.fasterxml.jackson.annotation.JsonInclude(
            com.fasterxml.jackson.annotation.JsonInclude.Include.NON_ABSENT)
        Optional<CommentAnchorInput> anchor) {
  /** Creates a hidden plain-text comment payload without rich-text metadata. */
  public static CommentInput hidden(TextSourceInput text, String author) {
    return new CommentInput(text, author, false, Optional.empty(), Optional.empty());
  }

  /** Creates a plain-text comment payload without rich-text runs or anchor overrides. */
  public static CommentInput plain(TextSourceInput text, String author, boolean visible) {
    return new CommentInput(text, author, visible, Optional.empty(), Optional.empty());
  }

  public CommentInput {
    text = requireNonBlankSource(text, "text");
    author = requireNonBlank(author, "author");
    Objects.requireNonNull(runs, "runs must not be null");
    Objects.requireNonNull(anchor, "anchor must not be null");
    if (runs.isPresent()) {
      java.util.List<RichTextRunInput> copiedRuns = java.util.List.copyOf(runs.orElseThrow());
      if (copiedRuns.isEmpty()) {
        throw new IllegalArgumentException("runs must not be empty when provided");
      }
      for (RichTextRunInput run : copiedRuns) {
        Objects.requireNonNull(run, "runs must not contain null values");
      }
      validateRunsAgainstInlineText(text, copiedRuns);
      runs = Optional.of(copiedRuns);
    }
  }

  /** Reads one comment payload with explicit visibility. */
  @JsonCreator
  public CommentInput(
      @JsonProperty("text") TextSourceInput text,
      @JsonProperty("author") String author,
      @JsonProperty("visible") Boolean visible,
      @JsonProperty("runs") Optional<java.util.List<RichTextRunInput>> runs,
      @JsonProperty("anchor") Optional<CommentAnchorInput> anchor) {
    this(
        text,
        author,
        Objects.requireNonNull(visible, "visible must not be null").booleanValue(),
        Objects.requireNonNull(runs, "runs must not be null"),
        Objects.requireNonNull(anchor, "anchor must not be null"));
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
