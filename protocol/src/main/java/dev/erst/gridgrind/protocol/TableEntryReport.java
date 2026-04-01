package dev.erst.gridgrind.protocol;

import java.util.List;
import java.util.Objects;

/** Protocol-facing factual report for one table loaded from a workbook. */
public record TableEntryReport(
    String name,
    String sheetName,
    String range,
    int headerRowCount,
    int totalsRowCount,
    List<String> columnNames,
    TableStyleReport style,
    boolean hasAutofilter) {
  public TableEntryReport {
    Objects.requireNonNull(name, "name must not be null");
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    if (sheetName.isBlank()) {
      throw new IllegalArgumentException("sheetName must not be blank");
    }
    Objects.requireNonNull(range, "range must not be null");
    if (headerRowCount < 0) {
      throw new IllegalArgumentException("headerRowCount must not be negative");
    }
    if (totalsRowCount < 0) {
      throw new IllegalArgumentException("totalsRowCount must not be negative");
    }
    Objects.requireNonNull(columnNames, "columnNames must not be null");
    columnNames = List.copyOf(columnNames);
    for (String columnName : columnNames) {
      Objects.requireNonNull(columnName, "columnNames must not contain nulls");
    }
    Objects.requireNonNull(style, "style must not be null");
  }
}
