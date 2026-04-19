package dev.erst.gridgrind.excel;

/** Engine-side reason emitted when a formula is not immediately evaluable. */
public enum ExcelFormulaCapabilityIssue {
  INVALID_FORMULA,
  MISSING_EXTERNAL_WORKBOOK,
  UNREGISTERED_USER_DEFINED_FUNCTION,
  UNSUPPORTED_FORMULA
}
