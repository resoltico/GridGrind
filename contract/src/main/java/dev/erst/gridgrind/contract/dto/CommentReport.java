package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonInclude;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import java.util.stream.Collectors;

/** Factual comment metadata returned for analyzed cells that carry a comment. */
public record CommentReport(
    String text,
    String author,
    boolean visible,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<List<RichTextRunReport>> runs,
    @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<CommentAnchorReport> anchor) {
  /** Creates a plain comment report without rich runs or anchor metadata. */
  public CommentReport(String text, String author, boolean visible) {
    this(text, author, visible, Optional.empty(), Optional.empty());
  }

  public CommentReport {
    Objects.requireNonNull(text, "text must not be null");
    Objects.requireNonNull(author, "author must not be null");
    runs = WorkbookResultSupport.copyOptionalValues(runs, "runs");
    anchor = Objects.requireNonNullElseGet(anchor, Optional::empty);
    if (runs.isPresent()) {
      List<RichTextRunReport> copiedRuns = runs.orElseThrow();
      if (!text.equals(
          copiedRuns.stream().map(RichTextRunReport::text).collect(Collectors.joining()))) {
        throw new IllegalArgumentException("comment runs must concatenate to the plain text");
      }
    }
  }
}
