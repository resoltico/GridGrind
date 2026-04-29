package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
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
        true,
        style,
        new TextSourceInput.Inline(""),
        false,
        false,
        false,
        "",
        "",
        "",
        java.util.List.of());
  }

  /** Creates a table payload with default metadata outside name, range, totals, and style. */
  public static TableInput create(
      String name, String sheetName, String range, Boolean showTotalsRow, TableStyleInput style) {
    return new TableInput(name, sheetName, range, showTotalsRow, style);
  }

  public TableInput {
    name = ProtocolDefinedNameValidation.validateName(name);
    ExcelSheetNames.requireValid(sheetName, "sheetName");
    Objects.requireNonNull(range, "range must not be null");
    if (range.isBlank()) {
      throw new IllegalArgumentException("range must not be blank");
    }
    Objects.requireNonNull(showTotalsRow, "showTotalsRow must not be null");
    Objects.requireNonNull(hasAutofilter, "hasAutofilter must not be null");
    Objects.requireNonNull(style, "style must not be null");
    Objects.requireNonNull(comment, "comment must not be null");
    Objects.requireNonNull(published, "published must not be null");
    Objects.requireNonNull(insertRow, "insertRow must not be null");
    Objects.requireNonNull(insertRowShift, "insertRowShift must not be null");
    Objects.requireNonNull(headerRowCellStyle, "headerRowCellStyle must not be null");
    Objects.requireNonNull(dataCellStyle, "dataCellStyle must not be null");
    Objects.requireNonNull(totalsRowCellStyle, "totalsRowCellStyle must not be null");
    columns = java.util.List.copyOf(Objects.requireNonNull(columns, "columns must not be null"));
    for (TableColumnInput column : columns) {
      Objects.requireNonNull(column, "columns must not contain null values");
    }
  }

  @JsonCreator(mode = JsonCreator.Mode.DELEGATING)
  static TableInput create(TableInputJson json) {
    return new TableInput(
        json.name(),
        json.sheetName(),
        json.range(),
        Boolean.TRUE.equals(json.showTotalsRow()),
        json.hasAutofilter() == null ? Boolean.TRUE : json.hasAutofilter(),
        json.style(),
        json.comment() == null ? new TextSourceInput.Inline("") : json.comment(),
        Boolean.TRUE.equals(json.published()),
        Boolean.TRUE.equals(json.insertRow()),
        Boolean.TRUE.equals(json.insertRowShift()),
        json.headerRowCellStyle() == null ? "" : json.headerRowCellStyle(),
        json.dataCellStyle() == null ? "" : json.dataCellStyle(),
        json.totalsRowCellStyle() == null ? "" : json.totalsRowCellStyle(),
        json.columns() == null ? java.util.List.of() : json.columns());
  }

  private record TableInputJson(
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
      java.util.List<TableColumnInput> columns) {}
}
