package dev.erst.gridgrind.core.vocabulary;

import dev.erst.gridgrind.core.limits.Limit;
import java.util.Optional;

/** Shared Excel-facing limits for sheet sizing and sheet-view state. */
@dev.erst.gridgrind.core.limits.EnforcesLimit({
  dev.erst.gridgrind.core.limits.Limit.ID_COLUMN_WIDTH,
  dev.erst.gridgrind.core.limits.Limit.ID_ROW_HEIGHT,
  dev.erst.gridgrind.core.limits.Limit.ID_SHEET_ZOOM_PERCENTAGE
})
public final class ExcelSheetLayoutLimits {

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
    if (defaultColumnWidth > Limit.MAX_DEFAULT_COLUMN_WIDTH) {
      throw new IllegalArgumentException(
          fieldName
              + " must not exceed "
              + Limit.MAX_DEFAULT_COLUMN_WIDTH
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
    if (widthCharacters > Limit.MAX_COLUMN_WIDTH_CHARACTERS) {
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
    if (heightPoints > Limit.MAX_ROW_HEIGHT_POINTS) {
      return Optional.of(ExcelRowHeightViolation.TOO_LARGE);
    }
    if (Math.round(heightPoints * 20.0d) <= 0L) {
      return Optional.of(ExcelRowHeightViolation.NOT_VISIBLE);
    }
    return Optional.empty();
  }

  /** Returns the zoom-range violation, if any. */
  public static Optional<ExcelZoomViolation> zoomViolation(int zoomPercent) {
    return zoomPercent < Limit.MIN_ZOOM_PERCENT || zoomPercent > Limit.MAX_ZOOM_PERCENT
        ? Optional.of(ExcelZoomViolation.OUT_OF_RANGE)
        : Optional.empty();
  }
}
