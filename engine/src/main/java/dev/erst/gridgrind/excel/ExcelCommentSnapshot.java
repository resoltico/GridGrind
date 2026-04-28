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

  /** Returns the full authoring view of this factual comment snapshot. */
  public ExcelComment toAuthoringComment() {
    return new ExcelComment(
        text,
        author,
        visible,
        runs == null
            ? null
            : new ExcelRichText(
                runs.runs().stream()
                    .map(
                        run ->
                            new ExcelRichTextRun(
                                run.text(),
                                new ExcelCellFont(
                                    run.font().bold(),
                                    run.font().italic(),
                                    run.font().fontName(),
                                    run.font().fontHeight(),
                                    run.font().fontColor() == null
                                        ? null
                                        : ExcelColorSupport.copyOf(run.font().fontColor()),
                                    run.font().underline(),
                                    run.font().strikeout())))
                    .toList()),
        anchor == null
            ? null
            : new ExcelCommentAnchor(
                anchor.firstColumn(), anchor.firstRow(), anchor.lastColumn(), anchor.lastRow()));
  }
}
