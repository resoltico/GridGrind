package dev.erst.gridgrind.excel;

import org.apache.poi.ss.usermodel.DataConsolidateFunction;

/** Product-owned pivot data-field aggregation functions supported by POI XSSF pivot tables. */
public enum ExcelPivotDataConsolidateFunction {
  SUM(DataConsolidateFunction.SUM),
  COUNT(DataConsolidateFunction.COUNT),
  COUNT_NUMS(DataConsolidateFunction.COUNT_NUMS),
  AVERAGE(DataConsolidateFunction.AVERAGE),
  MAX(DataConsolidateFunction.MAX),
  MIN(DataConsolidateFunction.MIN),
  PRODUCT(DataConsolidateFunction.PRODUCT),
  STD_DEV(DataConsolidateFunction.STD_DEV),
  STD_DEVP(DataConsolidateFunction.STD_DEVP),
  VAR(DataConsolidateFunction.VAR),
  VARP(DataConsolidateFunction.VARP);

  private final DataConsolidateFunction poiFunction;

  ExcelPivotDataConsolidateFunction(DataConsolidateFunction poiFunction) {
    this.poiFunction = poiFunction;
  }

  /** Returns the matching POI enum constant for authoritative pivot creation. */
  public DataConsolidateFunction toPoiFunction() {
    return poiFunction;
  }

  /** Maps one POI pivot-data aggregation function back into the GridGrind contract. */
  public static ExcelPivotDataConsolidateFunction fromPoiFunction(
      DataConsolidateFunction poiFunction) {
    for (ExcelPivotDataConsolidateFunction value : values()) {
      if (value.poiFunction == poiFunction) {
        return value;
      }
    }
    throw new IllegalArgumentException(
        "Unsupported pivot data consolidate function: " + poiFunction);
  }
}
