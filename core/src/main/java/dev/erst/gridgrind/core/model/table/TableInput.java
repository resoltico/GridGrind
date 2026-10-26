package dev.erst.gridgrind.core.model.table;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonProperty;
import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.metadata.ProtocolBooleanDefault;
import dev.erst.gridgrind.core.metadata.ProtocolField;
import dev.erst.gridgrind.core.source.TextSourceInput;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.TableName;
import java.util.Objects;

/** Protocol-facing table definition attached to one table authoring request. */
@CatalogType(
    id = "TableInput",
    summary = "Workbook-global table definition for one SET_TABLE request.",
    displayOrder = 74)
public record TableInput(
    @dev.erst.gridgrind.core.metadata.FieldRule(
            kind = dev.erst.gridgrind.core.metadata.FieldRule.Kind.LENGTH_RANGE,
            minimum = 1,
            maximum =
                dev.erst.gridgrind.core.vocabulary.ProtocolConstraintValues
                    .DEFINED_NAME_MAX_CODE_POINTS)
        @dev.erst.gridgrind.core.metadata.FieldRule(
            kind = dev.erst.gridgrind.core.metadata.FieldRule.Kind.STRING_PATTERN,
            text = dev.erst.gridgrind.core.vocabulary.ProtocolConstraintValues.DEFINED_NAME_PATTERN)
        TableName name,
    SheetName sheetName,
    RangeRef range,
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
      TableName name,
      SheetName sheetName,
      RangeRef range,
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
    java.util.Objects.requireNonNull(name, "name must not be null");
    java.util.Objects.requireNonNull(sheetName, "sheetName must not be null");
    Objects.requireNonNull(range, "range must not be null");
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
      @JsonProperty("name") TableName name,
      @JsonProperty("sheetName") SheetName sheetName,
      @JsonProperty("range") RangeRef range,
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
