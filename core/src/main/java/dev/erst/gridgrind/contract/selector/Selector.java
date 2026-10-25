package dev.erst.gridgrind.contract.selector;

/** Marker for one immutable workbook-target selector family. */
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
