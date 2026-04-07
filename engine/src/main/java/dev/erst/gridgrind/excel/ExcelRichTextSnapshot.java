package dev.erst.gridgrind.excel;

import java.util.List;
import java.util.Objects;

/** Immutable factual snapshot of rich-text runs stored in one string cell. */
public record ExcelRichTextSnapshot(List<ExcelRichTextRunSnapshot> runs) {
  public ExcelRichTextSnapshot {
    Objects.requireNonNull(runs, "runs must not be null");
    runs = List.copyOf(runs);
    if (runs.isEmpty()) {
      throw new IllegalArgumentException("runs must not be empty");
    }
    for (ExcelRichTextRunSnapshot run : runs) {
      Objects.requireNonNull(run, "runs must not contain nulls");
    }
  }

  /** Returns the plain string value represented by the stored run list. */
  public String plainText() {
    StringBuilder builder = new StringBuilder();
    for (ExcelRichTextRunSnapshot run : runs) {
      builder.append(run.text());
    }
    return builder.toString();
  }
}
