package dev.erst.gridgrind.excel;

import java.util.List;
import java.util.Objects;

/** Factual table metadata loaded from a workbook. */
public record ExcelTableSnapshot(
    String name,
    String sheetName,
    String range,
    int headerRowCount,
    int totalsRowCount,
    List<String> columnNames,
    List<ExcelTableColumnSnapshot> columns,
    ExcelTableStyleSnapshot style,
    boolean hasAutofilter,
    String comment,
    boolean published,
    boolean insertRow,
    boolean insertRowShift,
    String headerRowCellStyle,
    String dataCellStyle,
    String totalsRowCellStyle) {
  /** Creates a table snapshot with defaulted per-column metadata and optional flags. */
  public ExcelTableSnapshot(
      String name,
      String sheetName,
      String range,
      int headerRowCount,
      int totalsRowCount,
      List<String> columnNames,
      ExcelTableStyleSnapshot style,
      boolean hasAutofilter) {
    this(
        name,
        sheetName,
        range,
        headerRowCount,
        totalsRowCount,
        columnNames,
        columnNames.stream()
            .map(columnName -> new ExcelTableColumnSnapshot(0L, columnName, "", "", "", ""))
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

  public ExcelTableSnapshot {
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
    Objects.requireNonNull(columns, "columns must not be null");
    columns = List.copyOf(columns);
    for (ExcelTableColumnSnapshot column : columns) {
      Objects.requireNonNull(column, "columns must not contain nulls");
    }
    comment = comment == null ? "" : comment;
    headerRowCellStyle = headerRowCellStyle == null ? "" : headerRowCellStyle;
    dataCellStyle = dataCellStyle == null ? "" : dataCellStyle;
    totalsRowCellStyle = totalsRowCellStyle == null ? "" : totalsRowCellStyle;
    Objects.requireNonNull(style, "style must not be null");
  }
}
