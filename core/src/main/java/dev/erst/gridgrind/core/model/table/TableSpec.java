package dev.erst.gridgrind.core.model.table;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.metadata.ProtocolField;
import dev.erst.gridgrind.core.metadata.WireDefault;
import dev.erst.gridgrind.core.source.TextSourceSpec;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.TableName;
import java.util.Objects;

/** Protocol-facing table definition attached to one table authoring request. */
@CatalogType(
    id = "TableSpec",
    summary = "Workbook-global table definition for one SET_TABLE request.",
    displayOrder = 74)
public record TableSpec(
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
    @ProtocolField(optional = true) @WireDefault("false") boolean showTotalsRow,
    @ProtocolField(optional = true) @WireDefault("true") boolean hasAutofilter,
    TableStyleSpec style,
    @ProtocolField(optional = true) @WireDefault(factory = TextSourceSpec.class, method = "empty")
        TextSourceSpec comment,
    @ProtocolField(optional = true) @WireDefault("false") boolean published,
    @ProtocolField(optional = true) @WireDefault("false") boolean insertRow,
    @ProtocolField(optional = true) @WireDefault("false") boolean insertRowShift,
    @ProtocolField(optional = true) @WireDefault String headerRowCellStyle,
    @ProtocolField(optional = true) @WireDefault String dataCellStyle,
    @ProtocolField(optional = true) @WireDefault String totalsRowCellStyle,
    @ProtocolField(optional = true) @WireDefault("[]") java.util.List<TableColumnSpec> columns) {
  /** Creates a table payload with explicit standard metadata. */
  public static TableSpec withDefaultMetadata(
      TableName name,
      SheetName sheetName,
      RangeRef range,
      boolean showTotalsRow,
      TableStyleSpec style) {
    return new TableSpec(
        name,
        sheetName,
        range,
        showTotalsRow,
        true,
        style,
        new TextSourceSpec.Inline(""),
        false,
        false,
        false,
        "",
        "",
        "",
        java.util.List.of());
  }

  public TableSpec {
    java.util.Objects.requireNonNull(name, "name must not be null");
    java.util.Objects.requireNonNull(sheetName, "sheetName must not be null");
    Objects.requireNonNull(range, "range must not be null");
    Objects.requireNonNull(style, "style must not be null");
    Objects.requireNonNull(comment, "comment must not be null");
    Objects.requireNonNull(headerRowCellStyle, "headerRowCellStyle must not be null");
    Objects.requireNonNull(dataCellStyle, "dataCellStyle must not be null");
    Objects.requireNonNull(totalsRowCellStyle, "totalsRowCellStyle must not be null");
    columns = java.util.List.copyOf(Objects.requireNonNull(columns, "columns must not be null"));
    for (TableColumnSpec column : columns) {
      Objects.requireNonNull(column, "columns must not contain null values");
    }
  }
}
