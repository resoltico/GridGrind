package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelComparisonOperator;
import java.util.Locale;
import java.util.Objects;
import java.util.Optional;

/** Shared comparison-rule validation and formula normalization for Excel rule families. */
final class ExcelComparisonFormulaSupport {
  private ExcelComparisonFormulaSupport() {}

  /** Validates one comparison-rule payload before formula normalization is applied. */
  static void validateComparisonRule(
      ExcelComparisonOperator operator, String formula1, String formula2) {
    Objects.requireNonNull(operator, "operator must not be null");
    requireNonBlank(formula1, "formula1");
    if ((operator == ExcelComparisonOperator.BETWEEN
            || operator == ExcelComparisonOperator.NOT_BETWEEN)
        && (formula2 == null || formula2.isBlank())) {
      throw new IllegalArgumentException(
          "formula2 must not be blank for " + operator.name().toLowerCase(Locale.ROOT));
    }
    if (operator != ExcelComparisonOperator.BETWEEN
        && operator != ExcelComparisonOperator.NOT_BETWEEN
        && formula2 != null
        && !formula2.isBlank()) {
      throw new IllegalArgumentException(
          "formula2 must be omitted unless operator is BETWEEN or NOT_BETWEEN");
    }
  }

  /** Normalizes one formula-like string by trimming and removing a leading {@code =}. */
  static String normalizeFormula(String value, String fieldName) {
    String normalized = requireNonBlank(value, fieldName).trim();
    if (normalized.startsWith("=")) {
      normalized = normalized.substring(1);
    }
    if (normalized.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return normalized;
  }

  /** Normalizes the optional second comparison operand when the operator requires it. */
  static Optional<String> normalizeOptionalComparisonUpperBound(
      ExcelComparisonOperator operator, String formula2) {
    if (operator == ExcelComparisonOperator.BETWEEN
        || operator == ExcelComparisonOperator.NOT_BETWEEN) {
      return Optional.of(normalizeFormula(formula2, "formula2"));
    }
    return Optional.empty();
  }

  /** Requires one string field to be present and nonblank. */
  static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
