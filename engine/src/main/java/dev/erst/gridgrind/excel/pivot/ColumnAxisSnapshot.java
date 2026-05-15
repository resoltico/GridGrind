package dev.erst.gridgrind.excel.pivot;

import java.util.List;

/** Snapshot of pivot-table column-axis labels and values-axis placement. */
public record ColumnAxisSnapshot(
    List<ExcelPivotTableSnapshot.Field> columnLabels, boolean valuesAxisOnColumns) {
  public ColumnAxisSnapshot {
    columnLabels = List.copyOf(columnLabels);
  }
}
