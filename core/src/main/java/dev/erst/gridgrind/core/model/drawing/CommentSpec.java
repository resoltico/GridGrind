package dev.erst.gridgrind.core.model.drawing;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.metadata.ProtocolField;
import dev.erst.gridgrind.core.metadata.WireDefault;
import dev.erst.gridgrind.core.model.style.RichTextRunSpec;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import java.util.Objects;
import java.util.Optional;

/** Protocol-facing plain-text comment payload used by cell comment operations. */
@CatalogType(
    id = "CommentSpec",
    summary =
        "Comment payload attached to one cell. Comments can carry ordered rich-text runs and an explicit anchor box.",
    displayOrder = 40)
public record CommentSpec(
    TextSourceSpec text,
    String author,
    @ProtocolField(optional = true) @WireDefault("false") boolean visible,
    Optional<java.util.List<RichTextRunSpec>> runs,
    Optional<CommentAnchorSpec> anchor) {
  /** Creates a hidden plain-text comment payload without rich-text metadata. */
  public static CommentSpec hidden(TextSourceSpec text, String author) {
    return new CommentSpec(text, author, false, Optional.empty(), Optional.empty());
  }

  /** Creates a plain-text comment payload without rich-text runs or anchor overrides. */
  public static CommentSpec plain(TextSourceSpec text, String author, boolean visible) {
    return new CommentSpec(text, author, visible, Optional.empty(), Optional.empty());
  }

  public CommentSpec {
    text = requireNonBlankSource(text, "text");
    author = requireNonBlank(author, "author");
    Objects.requireNonNull(runs, "runs must not be null");
    Objects.requireNonNull(anchor, "anchor must not be null");
    if (runs.isPresent()) {
      java.util.List<RichTextRunSpec> copiedRuns = java.util.List.copyOf(runs.orElseThrow());
      if (copiedRuns.isEmpty()) {
        throw new IllegalArgumentException("runs must not be empty when provided");
      }
      for (RichTextRunSpec run : copiedRuns) {
        Objects.requireNonNull(run, "runs must not contain null values");
      }
      validateRunsAgainstInlineText(text, copiedRuns);
      runs = Optional.of(copiedRuns);
    }
  }

  private static void validateRunsAgainstInlineText(
      TextSourceSpec text, java.util.List<RichTextRunSpec> runs) {
    if (!(text instanceof TextSourceSpec.Inline inlineText)) {
      return;
    }
    if (runs.stream().anyMatch(run -> !(run.source() instanceof TextSourceSpec.Inline))) {
      return;
    }
    String joined =
        runs.stream()
            .map(RichTextRunSpec::source)
            .map(TextSourceSpec.Inline.class::cast)
            .map(TextSourceSpec.Inline::text)
            .collect(java.util.stream.Collectors.joining());
    if (!inlineText.text().equals(joined)) {
      throw new IllegalArgumentException("comment runs must concatenate to the plain text");
    }
  }

  private static TextSourceSpec requireNonBlankSource(TextSourceSpec value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value instanceof TextSourceSpec.Inline inline && inline.text().isBlank()) {
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
