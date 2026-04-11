package dev.erst.gridgrind.protocol.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.excel.ExcelComparisonOperator;
import java.util.List;
import java.util.Objects;

/** Protocol-facing data-validation rule families. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = DataValidationRuleInput.ExplicitList.class, name = "EXPLICIT_LIST"),
  @JsonSubTypes.Type(value = DataValidationRuleInput.FormulaList.class, name = "FORMULA_LIST"),
  @JsonSubTypes.Type(value = DataValidationRuleInput.WholeNumber.class, name = "WHOLE_NUMBER"),
  @JsonSubTypes.Type(value = DataValidationRuleInput.DecimalNumber.class, name = "DECIMAL_NUMBER"),
  @JsonSubTypes.Type(value = DataValidationRuleInput.DateRule.class, name = "DATE"),
  @JsonSubTypes.Type(value = DataValidationRuleInput.TimeRule.class, name = "TIME"),
  @JsonSubTypes.Type(value = DataValidationRuleInput.TextLength.class, name = "TEXT_LENGTH"),
  @JsonSubTypes.Type(value = DataValidationRuleInput.CustomFormula.class, name = "CUSTOM_FORMULA")
})
public sealed interface DataValidationRuleInput
    permits DataValidationRuleInput.ExplicitList,
        DataValidationRuleInput.FormulaList,
        DataValidationRuleInput.WholeNumber,
        DataValidationRuleInput.DecimalNumber,
        DataValidationRuleInput.DateRule,
        DataValidationRuleInput.TimeRule,
        DataValidationRuleInput.TextLength,
        DataValidationRuleInput.CustomFormula {

  /** Explicit-list validation rule. */
  record ExplicitList(List<String> values) implements DataValidationRuleInput {
    public ExplicitList {
      values = copyValues(values);
    }
  }

  /** Formula-driven list validation rule. */
  record FormulaList(String formula) implements DataValidationRuleInput {
    public FormulaList {
      formula = requireNonBlank(formula, "formula");
    }
  }

  /** Whole-number comparison validation rule. */
  record WholeNumber(ExcelComparisonOperator operator, String formula1, String formula2)
      implements DataValidationRuleInput {
    public WholeNumber {
      Objects.requireNonNull(operator, "operator must not be null");
      formula1 = requireNonBlank(formula1, "formula1");
    }
  }

  /** Decimal-number comparison validation rule. */
  record DecimalNumber(ExcelComparisonOperator operator, String formula1, String formula2)
      implements DataValidationRuleInput {
    public DecimalNumber {
      Objects.requireNonNull(operator, "operator must not be null");
      formula1 = requireNonBlank(formula1, "formula1");
    }
  }

  /** Date comparison validation rule. */
  record DateRule(ExcelComparisonOperator operator, String formula1, String formula2)
      implements DataValidationRuleInput {
    public DateRule {
      Objects.requireNonNull(operator, "operator must not be null");
      formula1 = requireNonBlank(formula1, "formula1");
    }
  }

  /** Time comparison validation rule. */
  record TimeRule(ExcelComparisonOperator operator, String formula1, String formula2)
      implements DataValidationRuleInput {
    public TimeRule {
      Objects.requireNonNull(operator, "operator must not be null");
      formula1 = requireNonBlank(formula1, "formula1");
    }
  }

  /** Text-length comparison validation rule. */
  record TextLength(ExcelComparisonOperator operator, String formula1, String formula2)
      implements DataValidationRuleInput {
    public TextLength {
      Objects.requireNonNull(operator, "operator must not be null");
      formula1 = requireNonBlank(formula1, "formula1");
    }
  }

  /** Custom formula validation rule. */
  record CustomFormula(String formula) implements DataValidationRuleInput {
    public CustomFormula {
      formula = requireNonBlank(formula, "formula");
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
