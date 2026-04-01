package dev.erst.gridgrind.excel;

import java.util.List;
import java.util.Locale;
import java.util.Objects;

/** Typed data-validation rule families supported by GridGrind. */
public sealed interface ExcelDataValidationRule
    permits ExcelDataValidationRule.ExplicitList,
        ExcelDataValidationRule.FormulaList,
        ExcelDataValidationRule.WholeNumber,
        ExcelDataValidationRule.DecimalNumber,
        ExcelDataValidationRule.DateRule,
        ExcelDataValidationRule.TimeRule,
        ExcelDataValidationRule.TextLength,
        ExcelDataValidationRule.CustomFormula {

  /** List-validation rule with explicit allowed values. */
  record ExplicitList(List<String> values) implements ExcelDataValidationRule {
    public ExplicitList {
      values = copyValues(values);
      if (values.isEmpty()) {
        throw new IllegalArgumentException("values must not be empty");
      }
    }
  }

  /** List-validation rule driven by a formula expression. */
  record FormulaList(String formula) implements ExcelDataValidationRule {
    public FormulaList {
      formula = normalizeFormula(formula, "formula");
    }
  }

  /** Whole-number validation driven by one comparison operator and one or two operands. */
  record WholeNumber(ExcelComparisonOperator operator, String formula1, String formula2)
      implements ExcelDataValidationRule {
    public WholeNumber {
      validateComparisonRule(operator, formula1, formula2);
      formula1 = normalizeFormula(formula1, "formula1");
      formula2 = normalizeOptionalComparisonUpperBound(operator, formula2);
    }
  }

  /** Decimal-number validation driven by one comparison operator and one or two operands. */
  record DecimalNumber(ExcelComparisonOperator operator, String formula1, String formula2)
      implements ExcelDataValidationRule {
    public DecimalNumber {
      validateComparisonRule(operator, formula1, formula2);
      formula1 = normalizeFormula(formula1, "formula1");
      formula2 = normalizeOptionalComparisonUpperBound(operator, formula2);
    }
  }

  /** Date validation driven by one comparison operator and one or two operands. */
  record DateRule(ExcelComparisonOperator operator, String formula1, String formula2)
      implements ExcelDataValidationRule {
    public DateRule {
      validateComparisonRule(operator, formula1, formula2);
      formula1 = normalizeFormula(formula1, "formula1");
      formula2 = normalizeOptionalComparisonUpperBound(operator, formula2);
    }
  }

  /** Time validation driven by one comparison operator and one or two operands. */
  record TimeRule(ExcelComparisonOperator operator, String formula1, String formula2)
      implements ExcelDataValidationRule {
    public TimeRule {
      validateComparisonRule(operator, formula1, formula2);
      formula1 = normalizeFormula(formula1, "formula1");
      formula2 = normalizeOptionalComparisonUpperBound(operator, formula2);
    }
  }

  /** Text-length validation driven by one comparison operator and one or two operands. */
  record TextLength(ExcelComparisonOperator operator, String formula1, String formula2)
      implements ExcelDataValidationRule {
    public TextLength {
      validateComparisonRule(operator, formula1, formula2);
      formula1 = normalizeFormula(formula1, "formula1");
      formula2 = normalizeOptionalComparisonUpperBound(operator, formula2);
    }
  }

  /** Custom formula-validation rule. */
  record CustomFormula(String formula) implements ExcelDataValidationRule {
    public CustomFormula {
      formula = normalizeFormula(formula, "formula");
    }
  }

  private static void validateComparisonRule(
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

  private static String normalizeOptionalComparisonUpperBound(
      ExcelComparisonOperator operator, String formula2) {
    if (operator == ExcelComparisonOperator.BETWEEN
        || operator == ExcelComparisonOperator.NOT_BETWEEN) {
      return normalizeFormula(formula2, "formula2");
    }
    return null;
  }

  private static List<String> copyValues(List<String> values) {
    Objects.requireNonNull(values, "values must not be null");
    List<String> copy = List.copyOf(values);
    for (String value : copy) {
      requireNonBlank(value, "values");
    }
    return copy;
  }

  private static String normalizeFormula(String value, String fieldName) {
    String normalized = requireNonBlank(value, fieldName).trim();
    if (normalized.startsWith("=")) {
      normalized = normalized.substring(1);
    }
    if (normalized.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return normalized;
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
