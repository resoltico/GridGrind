package dev.erst.gridgrind.contract.dto;

import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.excel.foundation.ExcelSheetNames;
import java.util.Objects;

/** Protocol-facing table definition attached to one table authoring request. */
public record TableInput(
    String name,
    String sheetName,
    String range,
    Boolean showTotalsRow,
    Boolean hasAutofilter,
    TableStyleInput style,
    TextSourceInput comment,
    Boolean published,
    Boolean insertRow,
    Boolean insertRowShift,
    String headerRowCellStyle,
    String dataCellStyle,
    String totalsRowCellStyle,
    java.util.List<TableColumnInput> columns) {
  /** Creates a table payload with default metadata outside name, range, totals, and style. */
  public TableInput(
      String name, String sheetName, String range, Boolean showTotalsRow, TableStyleInput style) {
    this(
        name,
        sheetName,
        range,
        showTotalsRow,
        null,
        style,
        null,
        null,
        null,
        null,
        null,
        null,
        null,
        null);
  }

  public TableInput {
    name = ProtocolDefinedNameValidation.validateName(name);
    ExcelSheetNames.requireValid(sheetName, "sheetName");
    Objects.requireNonNull(range, "range must not be null");
    if (range.isBlank()) {
      throw new IllegalArgumentException("range must not be blank");
    }
    showTotalsRow = Boolean.TRUE.equals(showTotalsRow);
    hasAutofilter = hasAutofilter == null ? Boolean.TRUE : hasAutofilter;
    Objects.requireNonNull(style, "style must not be null");
    comment = comment == null ? new TextSourceInput.Inline("") : comment;
    published = Boolean.TRUE.equals(published);
    insertRow = Boolean.TRUE.equals(insertRow);
    insertRowShift = Boolean.TRUE.equals(insertRowShift);
    headerRowCellStyle = headerRowCellStyle == null ? "" : headerRowCellStyle;
    dataCellStyle = dataCellStyle == null ? "" : dataCellStyle;
    totalsRowCellStyle = totalsRowCellStyle == null ? "" : totalsRowCellStyle;
    columns = columns == null ? java.util.List.of() : java.util.List.copyOf(columns);
    for (TableColumnInput column : columns) {
      Objects.requireNonNull(column, "columns must not contain null values");
    }
  }
}
