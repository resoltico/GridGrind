package dev.erst.gridgrind.excel;

import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Factual conditional-formatting rule metadata loaded from one workbook block. */
public sealed interface ExcelConditionalFormattingRuleSnapshot
    permits ExcelConditionalFormattingRuleSnapshot.FormulaRule,
        ExcelConditionalFormattingRuleSnapshot.CellValueRule,
        ExcelConditionalFormattingRuleSnapshot.ColorScaleRule,
        ExcelConditionalFormattingRuleSnapshot.DataBarRule,
        ExcelConditionalFormattingRuleSnapshot.IconSetRule,
        ExcelConditionalFormattingRuleSnapshot.UnsupportedRule {

  /** Persisted priority value loaded from the workbook. */
  int priority();

  /** Persisted stop-if-true flag loaded from the workbook. */
  boolean stopIfTrue();

  /** Formula-driven conditional-formatting rule with one differential-style payload. */
  record FormulaRule(
      int priority, boolean stopIfTrue, String formula, ExcelDifferentialStyleSnapshot style)
      implements ExcelConditionalFormattingRuleSnapshot {
    public FormulaRule {
      requirePriority(priority);
      formula = ExcelComparisonFormulaSupport.requireNonBlank(formula, "formula");
    }
  }

  /** Cell-value comparison rule with one or two operands and one differential-style payload. */
  record CellValueRule(
      int priority,
      boolean stopIfTrue,
      ExcelComparisonOperator operator,
      String formula1,
      String formula2,
      ExcelDifferentialStyleSnapshot style)
      implements ExcelConditionalFormattingRuleSnapshot {
    public CellValueRule {
      requirePriority(priority);
      Objects.requireNonNull(operator, "operator must not be null");
      formula1 = ExcelComparisonFormulaSupport.requireNonBlank(formula1, "formula1");
      if (formula2 != null && formula2.isBlank()) {
        throw new IllegalArgumentException("formula2 must not be blank");
      }
    }
  }

  /** Color-scale rule reported with thresholds and RGB color control points. */
  record ColorScaleRule(
      int priority,
      boolean stopIfTrue,
      List<ExcelConditionalFormattingThresholdSnapshot> thresholds,
      List<String> colors)
      implements ExcelConditionalFormattingRuleSnapshot {
    public ColorScaleRule {
      requirePriority(priority);
      thresholds = copyThresholds(thresholds, "thresholds");
      colors = copyColors(colors);
    }
  }

  /** Data-bar rule reported with thresholds, direction, widths, and RGB fill color. */
  record DataBarRule(
      int priority,
      boolean stopIfTrue,
      String color,
      boolean iconOnly,
      boolean leftToRight,
      int widthMin,
      int widthMax,
      ExcelConditionalFormattingThresholdSnapshot minThreshold,
      ExcelConditionalFormattingThresholdSnapshot maxThreshold)
      implements ExcelConditionalFormattingRuleSnapshot {
    public DataBarRule {
      requirePriority(priority);
      color = ExcelRgbColorSupport.normalizeRgbHex(color, "color");
      if (widthMin < 0) {
        throw new IllegalArgumentException("widthMin must not be negative");
      }
      if (widthMax < 0) {
        throw new IllegalArgumentException("widthMax must not be negative");
      }
      Objects.requireNonNull(minThreshold, "minThreshold must not be null");
      Objects.requireNonNull(maxThreshold, "maxThreshold must not be null");
    }
  }

  /** Icon-set rule reported with persisted icon set, direction flags, and thresholds. */
  record IconSetRule(
      int priority,
      boolean stopIfTrue,
      ExcelConditionalFormattingIconSet iconSet,
      boolean iconOnly,
      boolean reversed,
      List<ExcelConditionalFormattingThresholdSnapshot> thresholds)
      implements ExcelConditionalFormattingRuleSnapshot {
    public IconSetRule {
      requirePriority(priority);
      Objects.requireNonNull(iconSet, "iconSet must not be null");
      thresholds = copyThresholds(thresholds, "thresholds");
    }
  }

  /** Loaded rule family that GridGrind can detect but not yet model directly. */
  record UnsupportedRule(int priority, boolean stopIfTrue, String kind, String detail)
      implements ExcelConditionalFormattingRuleSnapshot {
    public UnsupportedRule {
      requirePriority(priority);
      kind = ExcelComparisonFormulaSupport.requireNonBlank(kind, "kind");
      detail = ExcelComparisonFormulaSupport.requireNonBlank(detail, "detail");
    }
  }

  private static void requirePriority(int priority) {
    if (priority < 0) {
      throw new IllegalArgumentException("priority must not be negative");
    }
  }

  private static List<ExcelConditionalFormattingThresholdSnapshot> copyThresholds(
      List<ExcelConditionalFormattingThresholdSnapshot> thresholds, String fieldName) {
    Objects.requireNonNull(thresholds, fieldName + " must not be null");
    List<ExcelConditionalFormattingThresholdSnapshot> copy = List.copyOf(thresholds);
    for (ExcelConditionalFormattingThresholdSnapshot threshold : copy) {
      Objects.requireNonNull(threshold, fieldName + " must not contain null values");
    }
    return copy;
  }

  private static List<String> copyColors(List<String> colors) {
    Objects.requireNonNull(colors, "colors must not be null");
    List<String> copy = new ArrayList<>(List.copyOf(colors));
    for (int index = 0; index < copy.size(); index++) {
      copy.set(index, ExcelRgbColorSupport.normalizeRgbHex(copy.get(index), "colors"));
    }
    return List.copyOf(copy);
  }
}
