package dev.erst.gridgrind.excel;

import java.util.List;
import java.util.Objects;

/** Authorable conditional-formatting rule families supported by GridGrind. */
public sealed interface ExcelConditionalFormattingRule
    permits ExcelConditionalFormattingRule.FormulaRule,
        ExcelConditionalFormattingRule.CellValueRule,
        ExcelConditionalFormattingRule.ColorScaleRule,
        ExcelConditionalFormattingRule.DataBarRule,
        ExcelConditionalFormattingRule.IconSetRule,
        ExcelConditionalFormattingRule.Top10Rule {

  /** Formula-driven conditional-formatting rule with one differential-style payload. */
  record FormulaRule(String formula, boolean stopIfTrue, ExcelDifferentialStyle style)
      implements ExcelConditionalFormattingRule {
    public FormulaRule {
      formula = ExcelComparisonFormulaSupport.normalizeFormula(formula, "formula");
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
    }
  }

  /** Color-scale rule with authored thresholds and control-point colors. */
  record ColorScaleRule(
      List<ExcelConditionalFormattingThreshold> thresholds,
      List<ExcelColor> colors,
      boolean stopIfTrue)
      implements ExcelConditionalFormattingRule {
    public ColorScaleRule {
      thresholds = copyThresholds(thresholds, "thresholds");
      colors = copyColors(colors, "colors");
      if (thresholds.size() < 2) {
        throw new IllegalArgumentException("thresholds must contain at least 2 control points");
      }
      if (thresholds.size() != colors.size()) {
        throw new IllegalArgumentException("thresholds and colors must have the same size");
      }
    }
  }

  /** Data-bar rule with authored thresholds, widths, and color payload. */
  record DataBarRule(
      ExcelColor color,
      boolean iconOnly,
      int widthMin,
      int widthMax,
      ExcelConditionalFormattingThreshold minThreshold,
      ExcelConditionalFormattingThreshold maxThreshold,
      boolean stopIfTrue)
      implements ExcelConditionalFormattingRule {
    public DataBarRule {
      Objects.requireNonNull(color, "color must not be null");
      Objects.requireNonNull(minThreshold, "minThreshold must not be null");
      Objects.requireNonNull(maxThreshold, "maxThreshold must not be null");
      if (widthMin < 0) {
        throw new IllegalArgumentException("widthMin must not be negative");
      }
      if (widthMax < 0) {
        throw new IllegalArgumentException("widthMax must not be negative");
      }
      if (widthMax < widthMin) {
        throw new IllegalArgumentException("widthMax must not be less than widthMin");
      }
    }
  }

  /** Icon-set rule with authored thresholds and icon family. */
  record IconSetRule(
      ExcelConditionalFormattingIconSet iconSet,
      boolean iconOnly,
      boolean reversed,
      List<ExcelConditionalFormattingThreshold> thresholds,
      boolean stopIfTrue)
      implements ExcelConditionalFormattingRule {
    public IconSetRule {
      Objects.requireNonNull(iconSet, "iconSet must not be null");
      thresholds = copyThresholds(thresholds, "thresholds");
      if (thresholds.size() != iconSet.thresholdCount()) {
        throw new IllegalArgumentException(
            "thresholds must contain exactly "
                + iconSet.thresholdCount()
                + " entries for "
                + iconSet);
      }
    }
  }

  /** Top-N or bottom-N rule with one differential-style payload. */
  record Top10Rule(
      int rank, boolean percent, boolean bottom, boolean stopIfTrue, ExcelDifferentialStyle style)
      implements ExcelConditionalFormattingRule {
    public Top10Rule {
      if (rank <= 0) {
        throw new IllegalArgumentException("rank must be greater than 0");
      }
    }
  }

  private static List<ExcelConditionalFormattingThreshold> copyThresholds(
      List<ExcelConditionalFormattingThreshold> thresholds, String fieldName) {
    List<ExcelConditionalFormattingThreshold> copy =
        List.copyOf(Objects.requireNonNull(thresholds, fieldName + " must not be null"));
    for (ExcelConditionalFormattingThreshold threshold : copy) {
      Objects.requireNonNull(threshold, fieldName + " must not contain null values");
    }
    return copy;
  }

  private static List<ExcelColor> copyColors(List<ExcelColor> colors, String fieldName) {
    List<ExcelColor> copy =
        List.copyOf(Objects.requireNonNull(colors, fieldName + " must not be null"));
    for (ExcelColor color : copy) {
      Objects.requireNonNull(color, fieldName + " must not contain null values");
    }
    return copy;
  }
}
