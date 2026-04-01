package dev.erst.gridgrind.excel;

import org.apache.poi.ss.usermodel.DataValidation;

/** Error-alert style used by Excel data-validation rules. */
public enum ExcelDataValidationErrorStyle {
  STOP,
  WARNING,
  INFORMATION;

  /** Returns the matching Apache POI error-style constant. */
  public int toPoiErrorStyle() {
    return switch (this) {
      case STOP -> DataValidation.ErrorStyle.STOP;
      case WARNING -> DataValidation.ErrorStyle.WARNING;
      case INFORMATION -> DataValidation.ErrorStyle.INFO;
    };
  }

  /** Converts one Apache POI error-style constant into the GridGrind enum. */
  public static ExcelDataValidationErrorStyle fromPoiErrorStyle(int errorStyle) {
    return switch (errorStyle) {
      case DataValidation.ErrorStyle.STOP -> STOP;
      case DataValidation.ErrorStyle.WARNING -> WARNING;
      case DataValidation.ErrorStyle.INFO -> INFORMATION;
      default ->
          throw new IllegalArgumentException("Unsupported Apache POI error style: " + errorStyle);
    };
  }
}
