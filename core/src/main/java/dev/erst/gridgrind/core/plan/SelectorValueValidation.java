package dev.erst.gridgrind.core.plan;

/** Shared scalar validation for selector families. */
final class SelectorValueValidation {
  private SelectorValueValidation() {}

  static String requireNonBlank(String value, String fieldName) {
    return SelectorTextValidation.requireNonBlank(value, fieldName);
  }

  static String requireDefinedName(String value, String fieldName) {
    return SelectorTextValidation.requireDefinedName(value, fieldName);
  }

  static String requirePivotTableName(String value, String fieldName) {
    return SelectorTextValidation.requirePivotTableName(value, fieldName);
  }

  static dev.erst.gridgrind.core.vocabulary.RangeRef requireOrderedRange(
      dev.erst.gridgrind.core.vocabulary.RangeRef value, String fieldName) {
    return SelectorTextValidation.requireOrderedRange(value, fieldName);
  }

  static int requirePositive(int value, String fieldName) {
    return SelectorNumberValidation.requirePositive(value, fieldName);
  }

  static int requireNonNegative(int value, String fieldName) {
    return SelectorNumberValidation.requireNonNegative(value, fieldName);
  }

  static int requireNonZero(int value, String fieldName) {
    return SelectorNumberValidation.requireNonZero(value, fieldName);
  }

  static int requireRowIndexWithinBounds(int value, String fieldName) {
    return SelectorNumberValidation.requireRowIndexWithinBounds(value, fieldName);
  }

  static int requireColumnIndexWithinBounds(int value, String fieldName) {
    return SelectorNumberValidation.requireColumnIndexWithinBounds(value, fieldName);
  }

  static void requireWindowSize(int rowCount, int columnCount) { // LIM-001
    SelectorNumberValidation.requireWindowSize(rowCount, columnCount);
  }

  static String prefixedValidationMessage(String fieldName, String message) {
    if (message == null || message.isBlank() || message.startsWith(fieldName + " ")) {
      return message;
    }
    return fieldName + " " + message;
  }
}
