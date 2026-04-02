package dev.erst.gridgrind.protocol;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.excel.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingRule;
import java.util.Objects;

/** Protocol-facing authored conditional-formatting rule families. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(
      value = ConditionalFormattingRuleInput.FormulaRule.class,
      name = "FORMULA_RULE"),
  @JsonSubTypes.Type(
      value = ConditionalFormattingRuleInput.CellValueRule.class,
      name = "CELL_VALUE_RULE")
})
public sealed interface ConditionalFormattingRuleInput
    permits ConditionalFormattingRuleInput.FormulaRule,
        ConditionalFormattingRuleInput.CellValueRule {

  /** Converts this transport model into the engine rule model. */
  ExcelConditionalFormattingRule toExcelConditionalFormattingRule();

  /** Formula-driven conditional-formatting rule with one differential-style payload. */
  record FormulaRule(String formula, boolean stopIfTrue, DifferentialStyleInput style)
      implements ConditionalFormattingRuleInput {
    public FormulaRule {
      Objects.requireNonNull(formula, "formula must not be null");
      if (formula.isBlank()) {
        throw new IllegalArgumentException("formula must not be blank");
      }
      Objects.requireNonNull(style, "style must not be null");
    }

    @Override
    public ExcelConditionalFormattingRule toExcelConditionalFormattingRule() {
      return new ExcelConditionalFormattingRule.FormulaRule(
          formula, stopIfTrue, style.toExcelDifferentialStyle());
    }
  }

  /** Cell-value comparison rule with one or two operands and one differential-style payload. */
  record CellValueRule(
      ExcelComparisonOperator operator,
      String formula1,
      String formula2,
      boolean stopIfTrue,
      DifferentialStyleInput style)
      implements ConditionalFormattingRuleInput {
    public CellValueRule {
      Objects.requireNonNull(operator, "operator must not be null");
      Objects.requireNonNull(formula1, "formula1 must not be null");
      if (formula1.isBlank()) {
        throw new IllegalArgumentException("formula1 must not be blank");
      }
      Objects.requireNonNull(style, "style must not be null");
    }

    @Override
    public ExcelConditionalFormattingRule toExcelConditionalFormattingRule() {
      return new ExcelConditionalFormattingRule.CellValueRule(
          operator, formula1, formula2, stopIfTrue, style.toExcelDifferentialStyle());
    }
  }
}
