package dev.erst.gridgrind.contract.dto;

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
    List<TableColumnReport> columns,
    TableStyleReport style,
    boolean hasAutofilter,
    String comment,
    boolean published,
    boolean insertRow,
    boolean insertRowShift,
    String headerRowCellStyle,
    String dataCellStyle,
    String totalsRowCellStyle) {
  /** Creates a table report with defaulted per-column metadata and optional flags. */
  public TableEntryReport(
      String name,
      String sheetName,
      String range,
      int headerRowCount,
      int totalsRowCount,
      List<String> columnNames,
      TableStyleReport style,
      boolean hasAutofilter) {
    this(
        name,
        sheetName,
        range,
        headerRowCount,
        totalsRowCount,
        columnNames,
        columnNames.stream()
            .map(columnName -> new TableColumnReport(0L, columnName, "", "", "", ""))
            .toList(),
        style,
        hasAutofilter,
        "",
        false,
        false,
        false,
        "",
        "",
        "");
  }

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
    columnNames = copyColumnNames(columnNames);
    Objects.requireNonNull(columns, "columns must not be null");
    columns = copyColumns(columns);
    comment = comment == null ? "" : comment;
    headerRowCellStyle = headerRowCellStyle == null ? "" : headerRowCellStyle;
    dataCellStyle = dataCellStyle == null ? "" : dataCellStyle;
    totalsRowCellStyle = totalsRowCellStyle == null ? "" : totalsRowCellStyle;
    Objects.requireNonNull(style, "style must not be null");
  }

  private static List<String> copyColumnNames(List<String> columnNames) {
    for (String columnName : columnNames) {
      Objects.requireNonNull(columnName, "columnNames must not contain nulls");
    }
    return List.copyOf(columnNames);
  }

  private static List<TableColumnReport> copyColumns(List<TableColumnReport> columns) {
    for (TableColumnReport column : columns) {
      Objects.requireNonNull(column, "columns must not contain nulls");
    }
    return List.copyOf(columns);
  }
}
