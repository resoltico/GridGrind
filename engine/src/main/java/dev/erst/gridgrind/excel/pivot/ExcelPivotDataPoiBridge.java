package dev.erst.gridgrind.excel.pivot;

import dev.erst.gridgrind.excel.foundation.ExcelPivotDataConsolidateFunction;
import org.apache.poi.ss.usermodel.DataConsolidateFunction;

/** Maps pivot data aggregation functions between GridGrind and Apache POI. */
@SuppressWarnings("PMD.CommentRequired")
public final class ExcelPivotDataPoiBridge {
  private ExcelPivotDataPoiBridge() {}

  public static DataConsolidateFunction toPoi(ExcelPivotDataConsolidateFunction function) {
    if (function == null) {
      throw new IllegalArgumentException("Unsupported GridGrind pivot data function: null");
    }
    return DataConsolidateFunction.valueOf(function.name());
  }

  public static ExcelPivotDataConsolidateFunction fromPoi(DataConsolidateFunction function) {
    if (function == null) {
      throw new IllegalArgumentException("Unsupported Apache POI pivot data function: null");
    }
    return ExcelPivotDataConsolidateFunction.valueOf(function.name());
  }
}
