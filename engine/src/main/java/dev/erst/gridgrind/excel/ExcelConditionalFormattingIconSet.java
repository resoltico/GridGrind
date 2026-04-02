package dev.erst.gridgrind.excel;

import org.apache.poi.ss.usermodel.IconMultiStateFormatting;

/** Supported icon-set identifiers reported for workbook-loaded conditional formatting. */
public enum ExcelConditionalFormattingIconSet {
  GYR_3_ARROW,
  GREY_3_ARROWS,
  GYR_3_FLAGS,
  GYR_3_TRAFFIC_LIGHTS,
  GYR_3_TRAFFIC_LIGHTS_BOX,
  GYR_3_SHAPES,
  GYR_3_SYMBOLS_CIRCLE,
  GYR_3_SYMBOLS,
  GYR_4_ARROWS,
  GREY_4_ARROWS,
  RB_4_TRAFFIC_LIGHTS,
  RATINGS_4,
  GYRB_4_TRAFFIC_LIGHTS,
  GYYYR_5_ARROWS,
  GREY_5_ARROWS,
  RATINGS_5,
  QUARTERS_5;

  /** Converts one POI icon-set identifier into the GridGrind enum. */
  public static ExcelConditionalFormattingIconSet fromPoi(
      IconMultiStateFormatting.IconSet iconSet) {
    return switch (iconSet) {
      case GYR_3_ARROW -> GYR_3_ARROW;
      case GREY_3_ARROWS -> GREY_3_ARROWS;
      case GYR_3_FLAGS -> GYR_3_FLAGS;
      case GYR_3_TRAFFIC_LIGHTS -> GYR_3_TRAFFIC_LIGHTS;
      case GYR_3_TRAFFIC_LIGHTS_BOX -> GYR_3_TRAFFIC_LIGHTS_BOX;
      case GYR_3_SHAPES -> GYR_3_SHAPES;
      case GYR_3_SYMBOLS_CIRCLE -> GYR_3_SYMBOLS_CIRCLE;
      case GYR_3_SYMBOLS -> GYR_3_SYMBOLS;
      case GYR_4_ARROWS -> GYR_4_ARROWS;
      case GREY_4_ARROWS -> GREY_4_ARROWS;
      case RB_4_TRAFFIC_LIGHTS -> RB_4_TRAFFIC_LIGHTS;
      case RATINGS_4 -> RATINGS_4;
      case GYRB_4_TRAFFIC_LIGHTS -> GYRB_4_TRAFFIC_LIGHTS;
      case GYYYR_5_ARROWS -> GYYYR_5_ARROWS;
      case GREY_5_ARROWS -> GREY_5_ARROWS;
      case RATINGS_5 -> RATINGS_5;
      case QUARTERS_5 -> QUARTERS_5;
    };
  }
}
