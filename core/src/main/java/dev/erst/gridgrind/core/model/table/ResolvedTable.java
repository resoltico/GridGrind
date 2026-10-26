package dev.erst.gridgrind.core.model.table;

import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.TableName;
import java.util.List;
import java.util.Objects;

/** Table definition whose comment source has been materialized to text. */
public record ResolvedTable(
    TableName name,
    SheetName sheetName,
    RangeRef range,
    boolean showTotalsRow,
    boolean hasAutofilter,
    TableStyleSpec style,
    String comment,
    boolean published,
    boolean insertRow,
    boolean insertRowShift,
    String headerRowCellStyle,
    String dataCellStyle,
    String totalsRowCellStyle,
    List<TableColumnSpec> columns) {
  /** Creates a table with standard metadata outside its identity, range, totals and style. */
  public ResolvedTable(
      TableName name,
      SheetName sheetName,
      RangeRef range,
      boolean showTotalsRow,
      TableStyleSpec style) {
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
        List.of());
  }

  public ResolvedTable {
    Objects.requireNonNull(name, "name must not be null");
    Objects.requireNonNull(sheetName, "sheetName must not be null");
    Objects.requireNonNull(range, "range must not be null");
    Objects.requireNonNull(style, "style must not be null");
    Objects.requireNonNull(comment, "comment must not be null");
    Objects.requireNonNull(headerRowCellStyle, "headerRowCellStyle must not be null");
    Objects.requireNonNull(dataCellStyle, "dataCellStyle must not be null");
    Objects.requireNonNull(totalsRowCellStyle, "totalsRowCellStyle must not be null");
    columns = List.copyOf(columns);
  }
}
