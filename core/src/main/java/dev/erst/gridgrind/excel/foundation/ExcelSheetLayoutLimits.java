package dev.erst.gridgrind.excel.foundation;

import java.util.Optional;

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
    columnWidthViolation(widthCharacters)
        .ifPresent(
            violation -> {
              throw new IllegalArgumentException(
                  ExcelSheetLayoutViolationMessages.columnWidth(
                      fieldName, violation, widthCharacters));
            });
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
    rowHeightViolation(heightPoints)
        .ifPresent(
            violation -> {
              throw new IllegalArgumentException(
                  ExcelSheetLayoutViolationMessages.rowHeight(fieldName, violation, heightPoints));
            });
  }

  /** Validates one authored worksheet zoom percentage. */
  public static void requireZoomPercent(int zoomPercent, String fieldName) {
    zoomViolation(zoomPercent)
        .ifPresent(
            violation -> {
              throw new IllegalArgumentException(
                  ExcelSheetLayoutViolationMessages.zoom(fieldName, zoomPercent));
            });
  }

  /** Returns the first authored column-width violation, if any. */
  public static Optional<ExcelColumnWidthViolation> columnWidthViolation(double widthCharacters) {
    if (!Double.isFinite(widthCharacters)) {
      return Optional.of(ExcelColumnWidthViolation.NON_FINITE);
    }
    if (widthCharacters <= 0.0d) {
      return Optional.of(ExcelColumnWidthViolation.NON_POSITIVE);
    }
    if (widthCharacters > MAX_COLUMN_WIDTH_CHARACTERS) {
      return Optional.of(ExcelColumnWidthViolation.TOO_LARGE);
    }
    if (Math.round(widthCharacters * 256.0d) <= 0) {
      return Optional.of(ExcelColumnWidthViolation.NOT_VISIBLE);
    }
    return Optional.empty();
  }

  /** Returns the first authored row-height violation, if any. */
  public static Optional<ExcelRowHeightViolation> rowHeightViolation(double heightPoints) {
    if (!Double.isFinite(heightPoints)) {
      return Optional.of(ExcelRowHeightViolation.NON_FINITE);
    }
    if (heightPoints <= 0.0d) {
      return Optional.of(ExcelRowHeightViolation.NON_POSITIVE);
    }
    if (heightPoints > MAX_ROW_HEIGHT_POINTS) {
      return Optional.of(ExcelRowHeightViolation.TOO_LARGE);
    }
    if (Math.round(heightPoints * 20.0d) <= 0L) {
      return Optional.of(ExcelRowHeightViolation.NOT_VISIBLE);
    }
    return Optional.empty();
  }

  /** Returns the zoom-range violation, if any. */
  public static Optional<ExcelZoomViolation> zoomViolation(int zoomPercent) {
    return zoomPercent < MIN_ZOOM_PERCENT || zoomPercent > MAX_ZOOM_PERCENT
        ? Optional.of(ExcelZoomViolation.OUT_OF_RANGE)
        : Optional.empty();
  }
}
