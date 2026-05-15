package dev.erst.gridgrind.excel.pivot;

import java.util.List;
import java.util.Locale;

/** Ordered source columns for one resolved pivot-table authoring surface. */
@SuppressWarnings("PMD.CommentRequired")
public record SourceColumns(List<SourceColumn> columns) {
  public SourceColumns {
    columns = List.copyOf(columns);
  }

  public int relativeIndex(String name) {
    String expected = name.toUpperCase(Locale.ROOT);
    for (SourceColumn column : columns) {
      if (column.name().toUpperCase(Locale.ROOT).equals(expected)) {
        return column.relativeIndex();
      }
    }
    throw new IllegalArgumentException("pivot source column not found: " + name);
  }
}
