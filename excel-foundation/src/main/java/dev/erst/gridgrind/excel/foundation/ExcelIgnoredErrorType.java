package dev.erst.gridgrind.excel.foundation;

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
  UNLOCKED_FORMULA
}
