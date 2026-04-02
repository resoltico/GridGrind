package dev.erst.gridgrind.excel;

import org.apache.poi.ss.usermodel.ConditionalFormattingThreshold;

/** Threshold families used by color-scale, data-bar, and icon-set rules. */
public enum ExcelConditionalFormattingThresholdType {
  NUMBER,
  MIN,
  MAX,
  PERCENT,
  PERCENTILE,
  UNALLOCATED,
  FORMULA;

  /** Converts one POI threshold type into the GridGrind enum. */
  public static ExcelConditionalFormattingThresholdType fromPoi(
      ConditionalFormattingThreshold.RangeType type) {
    return switch (type) {
      case NUMBER -> NUMBER;
      case MIN -> MIN;
      case MAX -> MAX;
      case PERCENT -> PERCENT;
      case PERCENTILE -> PERCENTILE;
      case UNALLOCATED -> UNALLOCATED;
      case FORMULA -> FORMULA;
    };
  }
}
