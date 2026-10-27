package dev.erst.gridgrind.engine.formulas;

import dev.erst.gridgrind.core.vocabulary.ExcelComparisonOperator;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.Locale;
import java.util.Objects;
import java.util.Optional;

/** Shared comparison-rule validation and formula normalization for Excel rule families. */
public final class ComparisonFormulas {
  private ComparisonFormulas() {}

  /** Validates one comparison-rule payload before formula normalization is applied. */
  public static void validateComparisonRule(
      ExcelComparisonOperator operator, String formula1, Optional<String> formula2) {
    Objects.requireNonNull(operator, "operator must not be null");
    Objects.requireNonNull(formula2, "formula2 must not be null");
    TextValidation.requireNonBlank(formula1, "formula1");
    if ((operator == ExcelComparisonOperator.BETWEEN
            || operator == ExcelComparisonOperator.NOT_BETWEEN)
        && (formula2.isEmpty() || formula2.orElseThrow().isBlank())) {
      throw new IllegalArgumentException(
          "formula2 must not be blank for " + operator.name().toLowerCase(Locale.ROOT));
    }
    if (operator != ExcelComparisonOperator.BETWEEN
        && operator != ExcelComparisonOperator.NOT_BETWEEN
        && formula2.isPresent()
        && !formula2.orElseThrow().isBlank()) {
      throw new IllegalArgumentException(
          "formula2 must be omitted unless operator is BETWEEN or NOT_BETWEEN");
    }
  }

  /** Normalizes one formula-like string by trimming and removing a leading {@code =}. */
  public static String normalizeFormula(String value, String fieldName) {
    String normalized = TextValidation.requireNonBlank(value, fieldName).trim();
    if (normalized.startsWith("=")) {
      normalized = normalized.substring(1);
    }
    if (normalized.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return normalized;
  }

  /** Normalizes each present formula without changing operand presence semantics. */
  public static Optional<String> normalizeOptionalFormula(
      Optional<String> value, String fieldName) {
    return value.map(formula -> normalizeFormula(formula, fieldName));
  }

  /** Normalizes the optional second comparison operand when the operator requires it. */
  public static Optional<String> normalizeOptionalComparisonUpperBound(
      ExcelComparisonOperator operator, Optional<String> formula2) {
    Objects.requireNonNull(formula2, "formula2 must not be null");
    if (operator == ExcelComparisonOperator.BETWEEN
        || operator == ExcelComparisonOperator.NOT_BETWEEN) {
      return Optional.of(normalizeFormula(formula2.orElseThrow(), "formula2"));
    }
    return Optional.empty();
  }
}
