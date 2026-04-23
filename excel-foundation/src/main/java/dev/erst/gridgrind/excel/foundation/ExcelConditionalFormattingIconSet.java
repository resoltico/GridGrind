package dev.erst.gridgrind.excel.foundation;

/** Supported icon-set identifiers reported for workbook-loaded conditional formatting. */
public enum ExcelConditionalFormattingIconSet {
  GYR_3_ARROW(3),
  GREY_3_ARROWS(3),
  GYR_3_FLAGS(3),
  GYR_3_TRAFFIC_LIGHTS(3),
  GYR_3_TRAFFIC_LIGHTS_BOX(3),
  GYR_3_SHAPES(3),
  GYR_3_SYMBOLS_CIRCLE(3),
  GYR_3_SYMBOLS(3),
  GYR_4_ARROWS(4),
  GREY_4_ARROWS(4),
  RB_4_TRAFFIC_LIGHTS(4),
  RATINGS_4(4),
  GYRB_4_TRAFFIC_LIGHTS(4),
  GYYYR_5_ARROWS(5),
  GREY_5_ARROWS(5),
  RATINGS_5(5),
  QUARTERS_5(5);

  private final int thresholdCount;

  ExcelConditionalFormattingIconSet(int thresholdCount) {
    this.thresholdCount = thresholdCount;
  }

  /** Returns the required threshold count for this icon-set family. */
  public int thresholdCount() {
    return thresholdCount;
  }
}
