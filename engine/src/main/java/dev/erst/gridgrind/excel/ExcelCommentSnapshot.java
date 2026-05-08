package dev.erst.gridgrind.excel;

import java.util.Objects;
import java.util.Optional;

/** Immutable factual comment metadata preserving plain text, runs, visibility, and anchor data. */
public record ExcelCommentSnapshot(
    String text,
    String author,
    boolean visible,
    Optional<ExcelRichTextSnapshot> runs,
    Optional<ExcelCommentAnchorSnapshot> anchor) {
  public ExcelCommentSnapshot {
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

  /** Returns the plain-text authoring view of this factual comment snapshot. */
  public ExcelComment toPlainComment() {
    return new ExcelComment(text, author, visible);
  }

  /** Returns the full authoring view of this factual comment snapshot. */
  public ExcelComment toAuthoringComment() {
    return new ExcelComment(
        text,
        author,
        visible,
        runs.map(
            snapshot ->
                new ExcelRichText(
                    snapshot.runs().stream()
                        .map(
                            run ->
                                new ExcelRichTextRun(
                                    run.text(),
                                    Optional.of(
                                        new ExcelCellFont(
                                            Optional.of(run.font().bold()),
                                            Optional.of(run.font().italic()),
                                            Optional.of(run.font().fontName()),
                                            Optional.of(run.font().fontHeight()),
                                            Optional.ofNullable(
                                                run.font().fontColor() == null
                                                    ? null
                                                    : ExcelColorSupport.copyOf(
                                                        run.font().fontColor())),
                                            Optional.of(run.font().underline()),
                                            Optional.of(run.font().strikeout())))))
                        .toList())),
        anchor.map(
            snapshot ->
                new ExcelCommentAnchor(
                    snapshot.firstColumn(),
                    snapshot.firstRow(),
                    snapshot.lastColumn(),
                    snapshot.lastRow())));
  }
}
