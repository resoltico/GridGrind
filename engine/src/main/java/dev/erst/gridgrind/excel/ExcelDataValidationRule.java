package dev.erst.gridgrind.excel;

import dev.erst.gridgrind.excel.foundation.ExcelComparisonOperator;
import java.util.List;
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
    }
  }

  /** List-validation rule driven by a formula expression. */
  record FormulaList(String formula) implements ExcelDataValidationRule {
    public FormulaList {
      formula = ExcelComparisonFormulaSupport.normalizeFormula(formula, "formula");
    }
  }

  /** Whole-number validation driven by one comparison operator and one or two operands. */
  record WholeNumber(ExcelComparisonOperator operator, String formula1, String formula2)
      implements ExcelDataValidationRule {
    public WholeNumber {
      ExcelComparisonFormulaSupport.validateComparisonRule(operator, formula1, formula2);
      formula1 = ExcelComparisonFormulaSupport.normalizeFormula(formula1, "formula1");
      formula2 =
          ExcelComparisonFormulaSupport.normalizeOptionalComparisonUpperBound(operator, formula2);
    }
  }

  /** Decimal-number validation driven by one comparison operator and one or two operands. */
  record DecimalNumber(ExcelComparisonOperator operator, String formula1, String formula2)
      implements ExcelDataValidationRule {
    public DecimalNumber {
      ExcelComparisonFormulaSupport.validateComparisonRule(operator, formula1, formula2);
      formula1 = ExcelComparisonFormulaSupport.normalizeFormula(formula1, "formula1");
      formula2 =
          ExcelComparisonFormulaSupport.normalizeOptionalComparisonUpperBound(operator, formula2);
    }
  }

  /** Date validation driven by one comparison operator and one or two operands. */
  record DateRule(ExcelComparisonOperator operator, String formula1, String formula2)
      implements ExcelDataValidationRule {
    public DateRule {
      ExcelComparisonFormulaSupport.validateComparisonRule(operator, formula1, formula2);
      formula1 = ExcelComparisonFormulaSupport.normalizeFormula(formula1, "formula1");
      formula2 =
          ExcelComparisonFormulaSupport.normalizeOptionalComparisonUpperBound(operator, formula2);
    }
  }

  /** Time validation driven by one comparison operator and one or two operands. */
  record TimeRule(ExcelComparisonOperator operator, String formula1, String formula2)
      implements ExcelDataValidationRule {
    public TimeRule {
      ExcelComparisonFormulaSupport.validateComparisonRule(operator, formula1, formula2);
      formula1 = ExcelComparisonFormulaSupport.normalizeFormula(formula1, "formula1");
      formula2 =
          ExcelComparisonFormulaSupport.normalizeOptionalComparisonUpperBound(operator, formula2);
    }
  }

  /** Text-length validation driven by one comparison operator and one or two operands. */
  record TextLength(ExcelComparisonOperator operator, String formula1, String formula2)
      implements ExcelDataValidationRule {
    public TextLength {
      ExcelComparisonFormulaSupport.validateComparisonRule(operator, formula1, formula2);
      formula1 = ExcelComparisonFormulaSupport.normalizeFormula(formula1, "formula1");
      formula2 =
          ExcelComparisonFormulaSupport.normalizeOptionalComparisonUpperBound(operator, formula2);
    }
  }

  /** Custom formula-validation rule. */
  record CustomFormula(String formula) implements ExcelDataValidationRule {
    public CustomFormula {
      formula = ExcelComparisonFormulaSupport.normalizeFormula(formula, "formula");
    }
  }

  private static List<String> copyValues(List<String> values) {
    Objects.requireNonNull(values, "values must not be null");
    List<String> copy = List.copyOf(values);
    for (String value : copy) {
      requireNonBlank(value, "values");
    }
    return copy;
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }
}
