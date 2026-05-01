package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.excel.foundation.ExcelSheetNames;
import java.util.Objects;

/** Protocol-facing table definition attached to one table authoring request. */
public record TableInput(
    String name,
    String sheetName,
    String range,
    boolean showTotalsRow,
    boolean hasAutofilter,
    TableStyleInput style,
    TextSourceInput comment,
    boolean published,
    boolean insertRow,
    boolean insertRowShift,
    String headerRowCellStyle,
    String dataCellStyle,
    String totalsRowCellStyle,
    java.util.List<TableColumnInput> columns) {
  /** Creates a table payload with explicit standard metadata. */
  public static TableInput withDefaultMetadata(
      String name, String sheetName, String range, boolean showTotalsRow, TableStyleInput style) {
    return new TableInput(
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

  public TableInput {
    name = ProtocolDefinedNameValidation.validateName(name);
    ExcelSheetNames.requireValid(sheetName, "sheetName");
    Objects.requireNonNull(range, "range must not be null");
    if (range.isBlank()) {
      throw new IllegalArgumentException("range must not be blank");
    }
    Objects.requireNonNull(style, "style must not be null");
    Objects.requireNonNull(comment, "comment must not be null");
    Objects.requireNonNull(headerRowCellStyle, "headerRowCellStyle must not be null");
    Objects.requireNonNull(dataCellStyle, "dataCellStyle must not be null");
    Objects.requireNonNull(totalsRowCellStyle, "totalsRowCellStyle must not be null");
    columns = java.util.List.copyOf(Objects.requireNonNull(columns, "columns must not be null"));
    for (TableColumnInput column : columns) {
      Objects.requireNonNull(column, "columns must not contain null values");
    }
  }

  /** Creates a table payload from the full authored wire shape. */
  @JsonCreator
  @SuppressWarnings("PMD.ExcessiveParameterList")
  public TableInput(
      @JsonProperty("name") String name,
      @JsonProperty("sheetName") String sheetName,
      @JsonProperty("range") String range,
      @JsonProperty("showTotalsRow") Boolean showTotalsRow,
      @JsonProperty("hasAutofilter") Boolean hasAutofilter,
      @JsonProperty("style") TableStyleInput style,
      @JsonProperty("comment") TextSourceInput comment,
      @JsonProperty("published") Boolean published,
      @JsonProperty("insertRow") Boolean insertRow,
      @JsonProperty("insertRowShift") Boolean insertRowShift,
      @JsonProperty("headerRowCellStyle") String headerRowCellStyle,
      @JsonProperty("dataCellStyle") String dataCellStyle,
      @JsonProperty("totalsRowCellStyle") String totalsRowCellStyle,
      @JsonProperty("columns") java.util.List<TableColumnInput> columns) {
    this(
        name,
        sheetName,
        range,
        Objects.requireNonNull(showTotalsRow, "showTotalsRow must not be null").booleanValue(),
        Objects.requireNonNull(hasAutofilter, "hasAutofilter must not be null").booleanValue(),
        Objects.requireNonNull(style, "style must not be null"),
        Objects.requireNonNull(comment, "comment must not be null"),
        Objects.requireNonNull(published, "published must not be null").booleanValue(),
        Objects.requireNonNull(insertRow, "insertRow must not be null").booleanValue(),
        Objects.requireNonNull(insertRowShift, "insertRowShift must not be null").booleanValue(),
        Objects.requireNonNull(headerRowCellStyle, "headerRowCellStyle must not be null"),
        Objects.requireNonNull(dataCellStyle, "dataCellStyle must not be null"),
        Objects.requireNonNull(totalsRowCellStyle, "totalsRowCellStyle must not be null"),
        Objects.requireNonNull(columns, "columns must not be null"));
  }
}
