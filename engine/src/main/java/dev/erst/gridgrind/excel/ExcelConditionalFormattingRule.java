package dev.erst.gridgrind.excel;

import java.util.Objects;

/** Authorable conditional-formatting rule families supported by GridGrind. */
public sealed interface ExcelConditionalFormattingRule
    permits ExcelConditionalFormattingRule.FormulaRule,
        ExcelConditionalFormattingRule.CellValueRule {

  /** Formula-driven conditional-formatting rule with one differential-style payload. */
  record FormulaRule(String formula, boolean stopIfTrue, ExcelDifferentialStyle style)
      implements ExcelConditionalFormattingRule {
    public FormulaRule {
      formula = ExcelComparisonFormulaSupport.normalizeFormula(formula, "formula");
      Objects.requireNonNull(style, "style must not be null");
    }
  }

  /** Cell-value comparison rule with one or two operands and one differential-style payload. */
  record CellValueRule(
      ExcelComparisonOperator operator,
      String formula1,
      String formula2,
      boolean stopIfTrue,
      ExcelDifferentialStyle style)
      implements ExcelConditionalFormattingRule {
    public CellValueRule {
      ExcelComparisonFormulaSupport.validateComparisonRule(operator, formula1, formula2);
      formula1 = ExcelComparisonFormulaSupport.normalizeFormula(formula1, "formula1");
      formula2 =
          ExcelComparisonFormulaSupport.normalizeOptionalComparisonUpperBound(operator, formula2);
      Objects.requireNonNull(style, "style must not be null");
    }
  }
}
