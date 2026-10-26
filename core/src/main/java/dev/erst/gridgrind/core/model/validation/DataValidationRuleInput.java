package dev.erst.gridgrind.core.model.validation;

import com.fasterxml.jackson.annotation.JsonInclude;
import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.model.formula.FormulaInputSecurity;
import dev.erst.gridgrind.core.vocabulary.ExcelComparisonOperator;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Protocol-facing data-validation rule families. */
@Union(discriminator = "type", catalog = true, displayOrder = 62)
public sealed interface DataValidationRuleInput
    permits dev.erst.gridgrind.core.model.validation.DataValidationRuleInput.ExplicitList,
        dev.erst.gridgrind.core.model.validation.DataValidationRuleInput.FormulaList,
        dev.erst.gridgrind.core.model.validation.DataValidationRuleInput.WholeNumber,
        dev.erst.gridgrind.core.model.validation.DataValidationRuleInput.DecimalNumber,
        dev.erst.gridgrind.core.model.validation.DataValidationRuleInput.DateRule,
        dev.erst.gridgrind.core.model.validation.DataValidationRuleInput.TimeRule,
        dev.erst.gridgrind.core.model.validation.DataValidationRuleInput.TextLength,
        dev.erst.gridgrind.core.model.validation.DataValidationRuleInput.CustomFormula {

  int MAX_EXPLICIT_LIST_FORMULA_LENGTH = 255; // LIM-040

  /** Explicit-list validation rule. */
  @Variant(
      id = "EXPLICIT_LIST",
      summary =
          "Allow only one of the supplied explicit values. An empty values array preserves Excel's explicit-empty-list state.")
  record ExplicitList(List<String> values) implements DataValidationRuleInput {
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
      implements DataValidationRuleInput {
    public FormulaList {
      formula = requireNonBlank(formula, "formula");
      FormulaInputSecurity.rejectDde(formula); // LIM-027
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
          @JsonInclude(JsonInclude.Include.NON_ABSENT)
          Optional<String> formula2)
      implements DataValidationRuleInput {
    public WholeNumber {
      Objects.requireNonNull(operator, "operator must not be null");
      formula1 = requireNonBlank(formula1, "formula1");
      FormulaInputSecurity.rejectDde(formula1); // LIM-027
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
          @JsonInclude(JsonInclude.Include.NON_ABSENT)
          Optional<String> formula2)
      implements DataValidationRuleInput {
    public DecimalNumber {
      Objects.requireNonNull(operator, "operator must not be null");
      formula1 = requireNonBlank(formula1, "formula1");
      FormulaInputSecurity.rejectDde(formula1); // LIM-027
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
          @JsonInclude(JsonInclude.Include.NON_ABSENT)
          Optional<String> formula2)
      implements DataValidationRuleInput {
    public DateRule {
      Objects.requireNonNull(operator, "operator must not be null");
      formula1 = requireNonBlank(formula1, "formula1");
      FormulaInputSecurity.rejectDde(formula1); // LIM-027
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
          @JsonInclude(JsonInclude.Include.NON_ABSENT)
          Optional<String> formula2)
      implements DataValidationRuleInput {
    public TimeRule {
      Objects.requireNonNull(operator, "operator must not be null");
      formula1 = requireNonBlank(formula1, "formula1");
      FormulaInputSecurity.rejectDde(formula1); // LIM-027
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
          @JsonInclude(JsonInclude.Include.NON_ABSENT)
          Optional<String> formula2)
      implements DataValidationRuleInput {
    public TextLength {
      Objects.requireNonNull(operator, "operator must not be null");
      formula1 = requireNonBlank(formula1, "formula1");
      FormulaInputSecurity.rejectDde(formula1); // LIM-027
      formula2 = normalizeOptionalComparisonUpperBound(operator, formula2);
    }
  }

  /** Custom formula validation rule. */
  @Variant(id = "CUSTOM_FORMULA", summary = "Allow values that satisfy a custom formula.")
  record CustomFormula(
      @dev.erst.gridgrind.core.metadata.FieldRule(
              kind = dev.erst.gridgrind.core.metadata.FieldRule.Kind.NON_BLANK)
          String formula)
      implements DataValidationRuleInput {
    public CustomFormula {
      formula = requireNonBlank(formula, "formula");
      FormulaInputSecurity.rejectDde(formula); // LIM-027
    }
  }

  private static List<String> copyValues(List<String> values) {
    Objects.requireNonNull(values, "values must not be null");
    List<String> copy = List.copyOf(values);
    for (String value : copy) {
      requireNonBlank(value, "values");
      if (value.indexOf(',') >= 0 || value.indexOf('"') >= 0) {
        throw new IllegalArgumentException(
            "values must not contain ',' or '\"' in an explicit list; use FORMULA_LIST for"
                + " values that require delimiters");
      }
    }
    int serializedLength = explicitListFormula(copy).length();
    if (serializedLength > MAX_EXPLICIT_LIST_FORMULA_LENGTH) {
      throw new IllegalArgumentException(
          "values must serialize to at most "
              + MAX_EXPLICIT_LIST_FORMULA_LENGTH
              + " characters for Excel data validation; got "
              + serializedLength);
    }
    return copy;
  }

  private static String explicitListFormula(List<String> values) {
    return "\"" + String.join(",", values) + "\"";
  }

  private static String requireNonBlank(String value, String fieldName) {
    Objects.requireNonNull(value, fieldName + " must not be null");
    if (value.isBlank()) {
      throw new IllegalArgumentException(fieldName + " must not be blank");
    }
    return value;
  }

  private static Optional<String> normalizeOptionalComparisonUpperBound(
      ExcelComparisonOperator operator, Optional<String> formula2) {
    Objects.requireNonNull(formula2, "formula2 must not be null");
    if (operator == ExcelComparisonOperator.BETWEEN
        || operator == ExcelComparisonOperator.NOT_BETWEEN) {
      String upperBound =
          requireNonBlank(
              formula2.orElseThrow(
                  () ->
                      new IllegalArgumentException(
                          "formula2 must not be blank for "
                              + operator.name().toLowerCase(java.util.Locale.ROOT))),
              "formula2");
      FormulaInputSecurity.rejectDde(upperBound); // LIM-027
      return Optional.of(upperBound);
    }
    if (formula2.isPresent()) {
      throw new IllegalArgumentException(
          "formula2 must be omitted unless operator is BETWEEN or NOT_BETWEEN");
    }
    return Optional.empty();
  }
}
