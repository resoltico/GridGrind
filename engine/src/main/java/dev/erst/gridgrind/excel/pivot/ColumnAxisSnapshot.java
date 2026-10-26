package dev.erst.gridgrind.excel.pivot;

import dev.erst.gridgrind.core.model.pivot.PivotTableReport;
import java.util.List;

/** Snapshot of pivot-table column-axis labels and values-axis placement. */
public record ColumnAxisSnapshot(
    List<PivotTableReport.Field> columnLabels, boolean valuesAxisOnColumns) {
  public ColumnAxisSnapshot {
    columnLabels = List.copyOf(columnLabels);
  }
}
