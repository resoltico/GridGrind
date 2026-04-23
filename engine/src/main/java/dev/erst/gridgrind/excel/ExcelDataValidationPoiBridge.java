package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelDataValidationErrorStyle;
import org.apache.poi.ss.usermodel.DataValidation;

/** Maps data-validation enums between GridGrind and Apache POI. */
final class ExcelDataValidationPoiBridge {
  private ExcelDataValidationPoiBridge() {}

  static int toPoiErrorStyle(ExcelDataValidationErrorStyle errorStyle) {
    return switch (errorStyle) {
      case STOP -> DataValidation.ErrorStyle.STOP;
      case WARNING -> DataValidation.ErrorStyle.WARNING;
      case INFORMATION -> DataValidation.ErrorStyle.INFO;
    };
  }

  static ExcelDataValidationErrorStyle fromPoiErrorStyle(int errorStyle) {
    return switch (errorStyle) {
      case DataValidation.ErrorStyle.STOP -> ExcelDataValidationErrorStyle.STOP;
      case DataValidation.ErrorStyle.WARNING -> ExcelDataValidationErrorStyle.WARNING;
      case DataValidation.ErrorStyle.INFO -> ExcelDataValidationErrorStyle.INFORMATION;
      default ->
          throw new IllegalArgumentException("Unsupported Apache POI error style: " + errorStyle);
    };
  }
}
