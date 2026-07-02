package dev.erst.gridgrind.excel.foundation;

/** Public wording for Excel sheet-layout validation failures. */
final class ExcelSheetLayoutViolationMessages {
  private ExcelSheetLayoutViolationMessages() {}

  static String columnWidth(
      String fieldName, ExcelColumnWidthViolation violation, double widthCharacters) {
    return switch (violation) {
      case NON_FINITE -> fieldName + " must be finite";
      case NON_POSITIVE -> fieldName + " must be greater than 0";
      case TOO_LARGE ->
          fieldName
              + " must not exceed "
              + ExcelSheetLayoutLimits.MAX_COLUMN_WIDTH_CHARACTERS
              + " (Excel column width limit): got "
              + widthCharacters;
      case NOT_VISIBLE ->
          fieldName
              + " is too small to produce a visible Excel column width: got "
              + widthCharacters;
    };
  }

  static String rowHeight(
      String fieldName, ExcelRowHeightViolation violation, double heightPoints) {
    return switch (violation) {
      case NON_FINITE -> fieldName + " must be finite";
      case NON_POSITIVE -> fieldName + " must be greater than 0";
      case TOO_LARGE ->
          fieldName
              + " must not exceed "
              + ExcelSheetLayoutLimits.MAX_ROW_HEIGHT_POINTS
              + " (Excel row height limit): got "
              + heightPoints;
      case NOT_VISIBLE ->
          fieldName + " is too small to produce a visible Excel row height: " + heightPoints;
    };
  }

  static String zoom(String fieldName, int zoomPercent) {
    return fieldName
        + " must be between "
        + ExcelSheetLayoutLimits.MIN_ZOOM_PERCENT
        + " and "
        + ExcelSheetLayoutLimits.MAX_ZOOM_PERCENT
        + " inclusive: "
        + zoomPercent;
  }
}
