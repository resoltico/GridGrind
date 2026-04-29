package dev.erst.gridgrind.contract.selector;

import tools.jackson.databind.annotation.JsonSerialize;

/** Marker for one immutable workbook-target selector family. */
@JsonSerialize(using = SelectorJsonSerializer.class)
public sealed interface Selector
    permits WorkbookSelector,
        SheetSelector,
        CellSelector,
        RangeSelector,
        RowBandSelector,
        ColumnBandSelector,
        DrawingObjectSelector,
        ChartSelector,
        TableSelector,
        PivotTableSelector,
        NamedRangeSelector,
        TableRowSelector,
        TableCellSelector {
  /** Returns the selector's declared target-count contract. */
  SelectorCardinality cardinality();
}
