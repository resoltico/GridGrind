package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelPivotDataConsolidateFunction;
import org.apache.poi.ss.usermodel.DataConsolidateFunction;

/** Maps pivot data aggregation functions between GridGrind and Apache POI. */
final class ExcelPivotDataPoiBridge {
  private ExcelPivotDataPoiBridge() {}

  static DataConsolidateFunction toPoi(ExcelPivotDataConsolidateFunction function) {
    if (function == null) {
      throw new IllegalArgumentException("Unsupported GridGrind pivot data function: null");
    }
    return switch (function) {
      case SUM -> DataConsolidateFunction.SUM;
      case COUNT -> DataConsolidateFunction.COUNT;
      case COUNT_NUMS -> DataConsolidateFunction.COUNT_NUMS;
      case AVERAGE -> DataConsolidateFunction.AVERAGE;
      case MAX -> DataConsolidateFunction.MAX;
      case MIN -> DataConsolidateFunction.MIN;
      case PRODUCT -> DataConsolidateFunction.PRODUCT;
      case STD_DEV -> DataConsolidateFunction.STD_DEV;
      case STD_DEVP -> DataConsolidateFunction.STD_DEVP;
      case VAR -> DataConsolidateFunction.VAR;
      case VARP -> DataConsolidateFunction.VARP;
    };
  }

  static ExcelPivotDataConsolidateFunction fromPoi(DataConsolidateFunction function) {
    if (function == null) {
      throw new IllegalArgumentException("Unsupported Apache POI pivot data function: null");
    }
    return switch (function) {
      case SUM -> ExcelPivotDataConsolidateFunction.SUM;
      case COUNT -> ExcelPivotDataConsolidateFunction.COUNT;
      case COUNT_NUMS -> ExcelPivotDataConsolidateFunction.COUNT_NUMS;
      case AVERAGE -> ExcelPivotDataConsolidateFunction.AVERAGE;
      case MAX -> ExcelPivotDataConsolidateFunction.MAX;
      case MIN -> ExcelPivotDataConsolidateFunction.MIN;
      case PRODUCT -> ExcelPivotDataConsolidateFunction.PRODUCT;
      case STD_DEV -> ExcelPivotDataConsolidateFunction.STD_DEV;
      case STD_DEVP -> ExcelPivotDataConsolidateFunction.STD_DEVP;
      case VAR -> ExcelPivotDataConsolidateFunction.VAR;
      case VARP -> ExcelPivotDataConsolidateFunction.VARP;
    };
  }
}
