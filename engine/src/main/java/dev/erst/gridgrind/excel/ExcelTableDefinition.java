package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Immutable workbook-core definition of one table to create or replace. */
public record ExcelTableDefinition(
    String name,
    String sheetName,
    String range,
    boolean showTotalsRow,
    boolean hasAutofilter,
    ExcelTableStyle style,
    String comment,
    boolean published,
    boolean insertRow,
    boolean insertRowShift,
    String headerRowCellStyle,
    String dataCellStyle,
    String totalsRowCellStyle,
    java.util.List<ExcelTableColumnDefinition> columns) {
  /** Creates a table definition with default metadata outside name, range, totals, and style. */
  public ExcelTableDefinition(
      String name, String sheetName, String range, boolean showTotalsRow, ExcelTableStyle style) {
    this(
        name,
        sheetName,
        range,
        showTotalsRow,
        true,
        style,
        "",
        false,
        false,
        false,
        "",
        "",
        "",
        java.util.List.of());
  }

  public ExcelTableDefinition {
    name = validateName(name);
    ExcelSheetNames.requireValid(sheetName, "sheetName");
    Objects.requireNonNull(range, "range must not be null");
    if (range.isBlank()) {
      throw new IllegalArgumentException("range must not be blank");
    }
    Objects.requireNonNull(style, "style must not be null");
    comment = comment == null ? "" : comment;
    headerRowCellStyle = headerRowCellStyle == null ? "" : headerRowCellStyle;
    dataCellStyle = dataCellStyle == null ? "" : dataCellStyle;
    totalsRowCellStyle = totalsRowCellStyle == null ? "" : totalsRowCellStyle;
    columns = columns == null ? java.util.List.of() : java.util.List.copyOf(columns);
    for (ExcelTableColumnDefinition column : columns) {
      Objects.requireNonNull(column, "columns must not contain null values");
    }
  }

  /** Validates and canonicalizes one table identifier. */
  public static String validateName(String name) {
    return ExcelNamedRangeDefinition.validateName(name);
  }
}
