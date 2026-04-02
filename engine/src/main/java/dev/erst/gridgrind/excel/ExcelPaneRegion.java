package dev.erst.gridgrind.excel;

import java.util.Objects;
import org.apache.poi.ss.usermodel.PaneType;

/** Identifies one visible sheet pane quadrant in workbook view state. */
public enum ExcelPaneRegion {
  UPPER_LEFT,
  UPPER_RIGHT,
  LOWER_LEFT,
  LOWER_RIGHT;

  /** Returns the GridGrind pane region corresponding to one POI pane type. */
  static ExcelPaneRegion fromPoi(PaneType paneType) {
    Objects.requireNonNull(paneType, "paneType must not be null");
    return switch (paneType) {
      case UPPER_LEFT -> UPPER_LEFT;
      case UPPER_RIGHT -> UPPER_RIGHT;
      case LOWER_LEFT -> LOWER_LEFT;
      case LOWER_RIGHT -> LOWER_RIGHT;
    };
  }

  /** Returns the POI pane type corresponding to this GridGrind pane region. */
  PaneType toPoi() {
    return switch (this) {
      case UPPER_LEFT -> PaneType.UPPER_LEFT;
      case UPPER_RIGHT -> PaneType.UPPER_RIGHT;
      case LOWER_LEFT -> PaneType.LOWER_LEFT;
      case LOWER_RIGHT -> PaneType.LOWER_RIGHT;
    };
  }
}
