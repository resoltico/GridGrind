package dev.erst.gridgrind.contract.dto;

import dev.erst.gridgrind.excel.foundation.ExcelReportedCellErrorLiteral;
import dev.erst.gridgrind.excel.foundation.ExcelStoredCellErrorLiteral;
import java.util.LinkedHashSet;
import java.util.Set;

/** Shared validation for stored-cell and reported-cell error literals on the protocol surface. */
final class CellErrorLiteralValidation {
  private CellErrorLiteralValidation() {}

  static String requireStoredErrorLiteral(String value, String fieldName) {
    String normalized = requireNonBlank(value, fieldName);
    if (ExcelStoredCellErrorLiteral.isSupportedWireValue(normalized)) {
      return ExcelStoredCellErrorLiteral.fromWireValue(normalized).wireValue();
    }
    if (ExcelReportedCellErrorLiteral.isSupportedWireValue(normalized)) {
      throw new IllegalArgumentException(
          fieldName
              + " must be one of "
              + String.join(", ", orderedStoredErrorLiterals())
              + "; computed evaluation-only states such as #CIRCULAR_REF! and"
              + " #FUNCTION_NOT_IMPLEMENTED! cannot be authored as stored cell values");
    }
    throw new IllegalArgumentException(
        fieldName
            + " must be one of "
            + String.join(", ", orderedStoredErrorLiterals())
            + "; for example #REF! or #DIV/0!");
  }

  static String requireReportedErrorLiteral(String value, String fieldName) {
    String normalized = requireNonBlank(value, fieldName);
    if (ExcelReportedCellErrorLiteral.isSupportedWireValue(normalized)) {
      return ExcelReportedCellErrorLiteral.fromWireValue(normalized).wireValue();
    }
    throw new IllegalArgumentException(
        fieldName
            + " must be one of "
            + String.join(", ", orderedReportedErrorLiterals())
            + "; for example #REF!, #DIV/0!, or #CIRCULAR_REF!");
  }

  static Set<String> orderedStoredErrorLiterals() {
    return new LinkedHashSet<>(ExcelStoredCellErrorLiteral.orderedWireValues());
  }

  static Set<String> orderedReportedErrorLiterals() {
    return new LinkedHashSet<>(ExcelReportedCellErrorLiteral.orderedWireValues());
  }

  private static String requireNonBlank(String value, String fieldName) {
    if (value == null) {
      throw new IllegalArgumentException(fieldName + " must not be null");
    }
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
