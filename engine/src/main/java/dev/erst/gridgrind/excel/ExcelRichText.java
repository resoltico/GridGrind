package dev.erst.gridgrind.excel;

import java.util.List;
import java.util.Objects;

/** Structured rich-text payload authored for one string cell value. */
public record ExcelRichText(List<ExcelRichTextRun> runs) {
  public ExcelRichText {
    Objects.requireNonNull(runs, "runs must not be null");
    runs = List.copyOf(runs);
    if (runs.isEmpty()) {
      throw new IllegalArgumentException("runs must not be empty");
    }
    for (ExcelRichTextRun run : runs) {
      Objects.requireNonNull(run, "runs must not contain nulls");
    }
  }

  /** Returns the plain string value represented by the ordered run list. */
  public String plainText() {
    StringBuilder builder = new StringBuilder();
    for (ExcelRichTextRun run : runs) {
      builder.append(run.text());
    }
    return builder.toString();
  }
}
