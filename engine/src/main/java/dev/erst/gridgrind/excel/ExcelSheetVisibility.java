package dev.erst.gridgrind.excel;

import java.util.Objects;
import org.apache.poi.ss.usermodel.SheetVisibility;

/** GridGrind-owned sheet-visibility states for workbook and sheet summaries and mutations. */
public enum ExcelSheetVisibility {
  VISIBLE,
  HIDDEN,
  VERY_HIDDEN;

  /** Converts one Apache POI visibility value into the GridGrind-owned visibility enum. */
  public static ExcelSheetVisibility fromPoi(SheetVisibility visibility) {
    Objects.requireNonNull(visibility, "visibility must not be null");
    return switch (visibility) {
      case VISIBLE -> VISIBLE;
      case HIDDEN -> HIDDEN;
      case VERY_HIDDEN -> VERY_HIDDEN;
    };
  }

  /** Converts this GridGrind-owned visibility enum into the Apache POI visibility enum. */
  public SheetVisibility toPoi() {
    return switch (this) {
      case VISIBLE -> SheetVisibility.VISIBLE;
      case HIDDEN -> SheetVisibility.HIDDEN;
      case VERY_HIDDEN -> SheetVisibility.VERY_HIDDEN;
    };
  }
}
