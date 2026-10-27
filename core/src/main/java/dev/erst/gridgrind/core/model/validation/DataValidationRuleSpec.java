package dev.erst.gridgrind.core.model.validation;

import dev.erst.gridgrind.core.limits.Limit;
import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.model.formula.FormulaInputSecurity;
import dev.erst.gridgrind.core.vocabulary.ExcelComparisonOperator;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Protocol-facing data-validation rule families. */
@Union(discriminator = "type", catalog = true, displayOrder = 62)
@dev.erst.gridgrind.core.limits.EnforcesLimit({
  dev.erst.gridgrind.core.limits.Limit.ID_DDE_REJECTION_IN_ALL_FORMULA_INPUTS,
  dev.erst.gridgrind.core.limits.Limit.ID_EXPLICIT_DATA_VALIDATION_LIST_FORMULA_LENGTH
})
public sealed interface DataValidationRuleSpec
    permits dev.erst.gridgrind.core.model.validation.DataValidationRuleSpec.ExplicitList,
        dev.erst.gridgrind.core.model.validation.DataValidationRuleSpec.FormulaList,
        dev.erst.gridgrind.core.model.validation.DataValidationRuleSpec.WholeNumber,
        dev.erst.gridgrind.core.model.validation.DataValidationRuleSpec.DecimalNumber,
        dev.erst.gridgrind.core.model.validation.DataValidationRuleSpec.DateRule,
        dev.erst.gridgrind.core.model.validation.DataValidationRuleSpec.TimeRule,
        dev.erst.gridgrind.core.model.validation.DataValidationRuleSpec.TextLength,
        dev.erst.gridgrind.core.model.validation.DataValidationRuleSpec.CustomFormula {

  /** Explicit-list validation rule. */
  @Variant(
      id = "EXPLICIT_LIST",
      summary =
          "Allow only one of the supplied explicit values. An empty values array preserves Excel's explicit-empty-list state.")
  record ExplicitList(List<String> values) implements DataValidationRuleSpec {
    public ExplicitList {
      values = copyValues(values);
    }
  }

  /** Formula-driven list validation rule. */
  @Variant(id = "FORMULA_LIST", summary = "Allow values from a formula-driven list expression.")
  record FormulaList(
      @dev.erst.gridgrind.core.metadata.FieldRule(
              kind = dev.erst.gridgrind.core.metadata.FieldRule.Kind.NON_BLANK)
          String formula)
      implements DataValidationRuleSpec {
    public FormulaList {
      formula = TextValidation.requireNonBlank(formula, "formula");
      FormulaInputSecurity.rejectDde(formula);
    }
  }

  /** Whole-number comparison validation rule. */
  @Variant(
      id = "WHOLE_NUMBER",
      summary =
          "Apply a whole-number comparison rule. formula2 is used only for BETWEEN and NOT_BETWEEN.")
  record WholeNumber(
      ExcelComparisonOperator operator,
      @dev.erst.gridgrind.core.metadata.FieldRule(
              kind = dev.erst.gridgrind.core.metadata.FieldRule.Kind.NON_BLANK)
          String formula1,
      @dev.erst.gridgrind.core.metadata.FieldRule(
              kind = dev.erst.gridgrind.core.metadata.FieldRule.Kind.NON_BLANK)
          Optional<String> formula2)
      implements DataValidationRuleSpec {
    public WholeNumber {
      Objects.requireNonNull(operator, "operator must not be null");
      formula1 = TextValidation.requireNonBlank(formula1, "formula1");
      FormulaInputSecurity.rejectDde(formula1);
      formula2 = normalizeOptionalComparisonUpperBound(operator, formula2);
    }
  }

  /** Decimal-number comparison validation rule. */
  @Variant(
      id = "DECIMAL_NUMBER",
      summary =
          "Apply a decimal-number comparison rule. formula2 is used only for BETWEEN and NOT_BETWEEN.")
  record DecimalNumber(
      ExcelComparisonOperator operator,
      @dev.erst.gridgrind.core.metadata.FieldRule(
              kind = dev.erst.gridgrind.core.metadata.FieldRule.Kind.NON_BLANK)
          String formula1,
      @dev.erst.gridgrind.core.metadata.FieldRule(
              kind = dev.erst.gridgrind.core.metadata.FieldRule.Kind.NON_BLANK)
          Optional<String> formula2)
      implements DataValidationRuleSpec {
    public DecimalNumber {
      Objects.requireNonNull(operator, "operator must not be null");
      formula1 = TextValidation.requireNonBlank(formula1, "formula1");
      FormulaInputSecurity.rejectDde(formula1);
      formula2 = normalizeOptionalComparisonUpperBound(operator, formula2);
    }
  }

  /** Date comparison validation rule. */
  @Variant(
      id = "DATE",
      summary = "Apply a date comparison rule. formula2 is used only for BETWEEN and NOT_BETWEEN.")
  record DateRule(
      ExcelComparisonOperator operator,
      @dev.erst.gridgrind.core.metadata.FieldRule(
              kind = dev.erst.gridgrind.core.metadata.FieldRule.Kind.NON_BLANK)
          String formula1,
      @dev.erst.gridgrind.core.metadata.FieldRule(
              kind = dev.erst.gridgrind.core.metadata.FieldRule.Kind.NON_BLANK)
          Optional<String> formula2)
      implements DataValidationRuleSpec {
    public DateRule {
      Objects.requireNonNull(operator, "operator must not be null");
      formula1 = TextValidation.requireNonBlank(formula1, "formula1");
      FormulaInputSecurity.rejectDde(formula1);
      formula2 = normalizeOptionalComparisonUpperBound(operator, formula2);
    }
  }

  /** Time comparison validation rule. */
  @Variant(
      id = "TIME",
      summary = "Apply a time comparison rule. formula2 is used only for BETWEEN and NOT_BETWEEN.")
  record TimeRule(
      ExcelComparisonOperator operator,
      @dev.erst.gridgrind.core.metadata.FieldRule(
              kind = dev.erst.gridgrind.core.metadata.FieldRule.Kind.NON_BLANK)
          String formula1,
      @dev.erst.gridgrind.core.metadata.FieldRule(
              kind = dev.erst.gridgrind.core.metadata.FieldRule.Kind.NON_BLANK)
          Optional<String> formula2)
      implements DataValidationRuleSpec {
    public TimeRule {
      Objects.requireNonNull(operator, "operator must not be null");
      formula1 = TextValidation.requireNonBlank(formula1, "formula1");
      FormulaInputSecurity.rejectDde(formula1);
      formula2 = normalizeOptionalComparisonUpperBound(operator, formula2);
    }
  }

  /** Text-length comparison validation rule. */
  @Variant(
      id = "TEXT_LENGTH",
      summary =
          "Apply a text-length comparison rule. formula2 is used only for BETWEEN and NOT_BETWEEN.")
  record TextLength(
      ExcelComparisonOperator operator,
      @dev.erst.gridgrind.core.metadata.FieldRule(
              kind = dev.erst.gridgrind.core.metadata.FieldRule.Kind.NON_BLANK)
          String formula1,
      @dev.erst.gridgrind.core.metadata.FieldRule(
              kind = dev.erst.gridgrind.core.metadata.FieldRule.Kind.NON_BLANK)
          Optional<String> formula2)
      implements DataValidationRuleSpec {
    public TextLength {
      Objects.requireNonNull(operator, "operator must not be null");
      formula1 = TextValidation.requireNonBlank(formula1, "formula1");
      FormulaInputSecurity.rejectDde(formula1);
      formula2 = normalizeOptionalComparisonUpperBound(operator, formula2);
    }
  }

  /** Custom formula validation rule. */
  @Variant(id = "CUSTOM_FORMULA", summary = "Allow values that satisfy a custom formula.")
  record CustomFormula(
      @dev.erst.gridgrind.core.metadata.FieldRule(
              kind = dev.erst.gridgrind.core.metadata.FieldRule.Kind.NON_BLANK)
          String formula)
      implements DataValidationRuleSpec {
    public CustomFormula {
      formula = TextValidation.requireNonBlank(formula, "formula");
      FormulaInputSecurity.rejectDde(formula);
    }
  }

  private static List<String> copyValues(List<String> values) {
    Objects.requireNonNull(values, "values must not be null");
    List<String> copy = List.copyOf(values);
    for (String value : copy) {
      TextValidation.requireNonBlank(value, "values");
      if (value.indexOf(',') >= 0 || value.indexOf('"') >= 0) {
        throw new IllegalArgumentException(
            "values must not contain ',' or '\"' in an explicit list; use FORMULA_LIST for"
                + " values that require delimiters");
      }
    }
    int serializedLength = explicitListFormula(copy).length();
    if (serializedLength > Limit.MAX_EXPLICIT_LIST_FORMULA_LENGTH) {
      throw new IllegalArgumentException(
          "values must serialize to at most "
              + Limit.MAX_EXPLICIT_LIST_FORMULA_LENGTH
              + " characters for Excel data validation; got "
              + serializedLength);
    }
    return copy;
  }

  private static String explicitListFormula(List<String> values) {
    return "\"" + String.join(",", values) + "\"";
  }

  private static Optional<String> normalizeOptionalComparisonUpperBound(
      ExcelComparisonOperator operator, Optional<String> formula2) {
    Objects.requireNonNull(formula2, "formula2 must not be null");
    if (operator == ExcelComparisonOperator.BETWEEN
        || operator == ExcelComparisonOperator.NOT_BETWEEN) {
      String upperBound =
          TextValidation.requireNonBlank(
              formula2.orElseThrow(
                  () ->
                      new IllegalArgumentException(
                          "formula2 must not be blank for "
                              + operator.name().toLowerCase(java.util.Locale.ROOT))),
              "formula2");
      FormulaInputSecurity.rejectDde(upperBound);
      return Optional.of(upperBound);
    }
    if (formula2.isPresent()) {
      throw new IllegalArgumentException(
          "formula2 must be omitted unless operator is BETWEEN or NOT_BETWEEN");
    }
    return Optional.empty();
  }
}
