package dev.erst.gridgrind.excel;

import org.apache.poi.ss.usermodel.IgnoredErrorType;

/** Supported ignored-error families exposed in sheet presentation state. */
public enum ExcelIgnoredErrorType {
  CALCULATED_COLUMN,
  EMPTY_CELL_REFERENCE,
  EVALUATION_ERROR,
  FORMULA,
  FORMULA_RANGE,
  LIST_DATA_VALIDATION,
  NUMBER_STORED_AS_TEXT,
  TWO_DIGIT_TEXT_YEAR,
  UNLOCKED_FORMULA;

  /** Converts this GridGrind-facing ignored-error family to the matching POI enum. */
  IgnoredErrorType toPoi() {
    return switch (this) {
      case CALCULATED_COLUMN -> IgnoredErrorType.CALCULATED_COLUMN;
      case EMPTY_CELL_REFERENCE -> IgnoredErrorType.EMPTY_CELL_REFERENCE;
      case EVALUATION_ERROR -> IgnoredErrorType.EVALUATION_ERROR;
      case FORMULA -> IgnoredErrorType.FORMULA;
      case FORMULA_RANGE -> IgnoredErrorType.FORMULA_RANGE;
      case LIST_DATA_VALIDATION -> IgnoredErrorType.LIST_DATA_VALIDATION;
      case NUMBER_STORED_AS_TEXT -> IgnoredErrorType.NUMBER_STORED_AS_TEXT;
      case TWO_DIGIT_TEXT_YEAR -> IgnoredErrorType.TWO_DIGIT_TEXT_YEAR;
      case UNLOCKED_FORMULA -> IgnoredErrorType.UNLOCKED_FORMULA;
    };
  }

  /** Converts one POI ignored-error family to the matching GridGrind-facing enum. */
  static ExcelIgnoredErrorType fromPoi(IgnoredErrorType ignoredErrorType) {
    return switch (ignoredErrorType) {
      case CALCULATED_COLUMN -> CALCULATED_COLUMN;
      case EMPTY_CELL_REFERENCE -> EMPTY_CELL_REFERENCE;
      case EVALUATION_ERROR -> EVALUATION_ERROR;
      case FORMULA -> FORMULA;
      case FORMULA_RANGE -> FORMULA_RANGE;
      case LIST_DATA_VALIDATION -> LIST_DATA_VALIDATION;
      case NUMBER_STORED_AS_TEXT -> NUMBER_STORED_AS_TEXT;
      case TWO_DIGIT_TEXT_YEAR -> TWO_DIGIT_TEXT_YEAR;
      case UNLOCKED_FORMULA -> UNLOCKED_FORMULA;
    };
  }
}
