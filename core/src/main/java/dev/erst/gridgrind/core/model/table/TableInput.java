package dev.erst.gridgrind.core.model.table;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import dev.erst.gridgrind.core.metadata.ProtocolBooleanDefault;
import dev.erst.gridgrind.core.metadata.ProtocolField;
import dev.erst.gridgrind.core.source.TextSourceInput;
import dev.erst.gridgrind.core.vocabulary.ProtocolDefinedNameValidation;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import java.util.Objects;

/** Protocol-facing table definition attached to one table authoring request. */
public record TableInput(
    String name,
    SheetName sheetName,
    String range,
    @ProtocolField(optional = true, booleanDefault = ProtocolBooleanDefault.FALSE)
        boolean showTotalsRow,
    @ProtocolField(optional = true, booleanDefault = ProtocolBooleanDefault.TRUE)
        boolean hasAutofilter,
    TableStyleInput style,
    @ProtocolField(optional = true) TextSourceInput comment,
    @ProtocolField(optional = true, booleanDefault = ProtocolBooleanDefault.FALSE)
        boolean published,
    @ProtocolField(optional = true, booleanDefault = ProtocolBooleanDefault.FALSE)
        boolean insertRow,
    @ProtocolField(optional = true, booleanDefault = ProtocolBooleanDefault.FALSE)
        boolean insertRowShift,
    @ProtocolField(optional = true) String headerRowCellStyle,
    @ProtocolField(optional = true) String dataCellStyle,
    @ProtocolField(optional = true) String totalsRowCellStyle,
    @ProtocolField(optional = true) java.util.List<TableColumnInput> columns) {
  /** Creates a table payload with explicit standard metadata. */
  public static TableInput withDefaultMetadata(
      String name,
      SheetName sheetName,
      String range,
      boolean showTotalsRow,
      TableStyleInput style) {
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
    java.util.Objects.requireNonNull(sheetName, "sheetName must not be null");
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
      @JsonProperty("sheetName") SheetName sheetName,
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
        ProtocolBooleanDefault.FALSE.resolve(showTotalsRow),
        ProtocolBooleanDefault.TRUE.resolve(hasAutofilter),
        Objects.requireNonNull(style, "style must not be null"),
        comment == null ? new TextSourceInput.Inline("") : comment,
        ProtocolBooleanDefault.FALSE.resolve(published),
        ProtocolBooleanDefault.FALSE.resolve(insertRow),
        ProtocolBooleanDefault.FALSE.resolve(insertRowShift),
        headerRowCellStyle == null ? "" : headerRowCellStyle,
        dataCellStyle == null ? "" : dataCellStyle,
        totalsRowCellStyle == null ? "" : totalsRowCellStyle,
        columns == null ? java.util.List.of() : columns);
  }
}
