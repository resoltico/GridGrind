package dev.erst.gridgrind.core.model.style;

import dev.erst.gridgrind.core.model.value.ResolvedCellWrite;
import java.util.List;
import java.util.Objects;

/** Materialized ordered rich-text runs ready for workbook authoring. */
public record ResolvedRichText(List<ResolvedRichTextRun> runs) implements ResolvedCellWrite {
  public ResolvedRichText {
    Objects.requireNonNull(runs, "runs must not be null");
    runs = List.copyOf(runs);
    if (runs.isEmpty()) {
      throw new IllegalArgumentException("runs must not be empty");
    }
    long length = 0;
    for (ResolvedRichTextRun run : runs) {
      Objects.requireNonNull(run, "runs must not contain nulls");
      length += run.text().length();
    }
    dev.erst.gridgrind.core.model.value.CellTextLimits.requireSupportedLength(length, "runs");
  }

  /** Returns the plain string value represented by the ordered run list. */
  public String plainText() {
    StringBuilder builder = new StringBuilder();
    for (ResolvedRichTextRun run : runs) {
      builder.append(run.text());
    }
    return builder.toString();
  }
}
