package dev.erst.gridgrind.excel.foundation;

/** Shared Excel-facing limits for sheet sizing and sheet-view state. */
public final class ExcelSheetLayoutLimits {
  public static final double MAX_COLUMN_WIDTH_CHARACTERS = 255.0d; // LIM-004
  public static final int MAX_DEFAULT_COLUMN_WIDTH = 255; // LIM-004
  public static final double MAX_ROW_HEIGHT_POINTS = 409.0d; // LIM-005
  public static final int MIN_ZOOM_PERCENT = 10; // LIM-022
  public static final int MAX_ZOOM_PERCENT = 400; // LIM-022

  private ExcelSheetLayoutLimits() {}

  /** Validates one authored column width expressed in Excel character units. */
  public static void requireColumnWidthCharacters(double widthCharacters, String fieldName) {
    requireFinitePositive(widthCharacters, fieldName);
    if (widthCharacters > MAX_COLUMN_WIDTH_CHARACTERS) {
      throw new IllegalArgumentException(
          fieldName
              + " must not exceed "
              + MAX_COLUMN_WIDTH_CHARACTERS
              + " (Excel column width limit): got "
              + widthCharacters);
    }
    if (Math.round(widthCharacters * 256.0d) <= 0) {
      throw new IllegalArgumentException(
          fieldName
              + " is too small to produce a visible Excel column width: got "
              + widthCharacters);
    }
  }

  /** Validates one authored default sheet column width expressed in whole Excel characters. */
  public static void requireDefaultColumnWidth(int defaultColumnWidth, String fieldName) {
    if (defaultColumnWidth <= 0) {
      throw new IllegalArgumentException(fieldName + " must be greater than 0");
    }
    if (defaultColumnWidth > MAX_DEFAULT_COLUMN_WIDTH) {
      throw new IllegalArgumentException(
          fieldName
              + " must not exceed "
              + MAX_DEFAULT_COLUMN_WIDTH
              + " (Excel column width limit): got "
              + defaultColumnWidth);
    }
  }

  /** Validates one authored row height expressed in Excel point units. */
  public static void requireRowHeightPoints(double heightPoints, String fieldName) {
    requireFinitePositive(heightPoints, fieldName);
    if (heightPoints > MAX_ROW_HEIGHT_POINTS) {
      throw new IllegalArgumentException(
          fieldName
              + " must not exceed "
              + MAX_ROW_HEIGHT_POINTS
              + " (Excel row height limit): got "
              + heightPoints);
    }
    if (Math.round(heightPoints * 20.0d) <= 0L) {
      throw new IllegalArgumentException(
          fieldName + " is too small to produce a visible Excel row height: " + heightPoints);
    }
  }

  /** Validates one authored worksheet zoom percentage. */
  public static void requireZoomPercent(int zoomPercent, String fieldName) {
    if (zoomPercent < MIN_ZOOM_PERCENT || zoomPercent > MAX_ZOOM_PERCENT) {
      throw new IllegalArgumentException(
          fieldName
              + " must be between "
              + MIN_ZOOM_PERCENT
              + " and "
              + MAX_ZOOM_PERCENT
              + " inclusive: "
              + zoomPercent);
    }
  }

  private static void requireFinitePositive(double value, String fieldName) {
    if (!Double.isFinite(value)) {
      throw new IllegalArgumentException(fieldName + " must be finite");
    }
    if (value <= 0.0d) {
      throw new IllegalArgumentException(fieldName + " must be greater than 0");
    }
  }
}
