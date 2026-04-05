package dev.erst.gridgrind.protocol.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.excel.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingIconSet;
import java.util.List;
import java.util.Objects;

/** Protocol-facing factual report for one conditional-formatting rule loaded from a workbook. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(
      value = ConditionalFormattingRuleReport.FormulaRule.class,
      name = "FORMULA_RULE"),
  @JsonSubTypes.Type(
      value = ConditionalFormattingRuleReport.CellValueRule.class,
      name = "CELL_VALUE_RULE"),
  @JsonSubTypes.Type(
      value = ConditionalFormattingRuleReport.ColorScaleRule.class,
      name = "COLOR_SCALE_RULE"),
  @JsonSubTypes.Type(
      value = ConditionalFormattingRuleReport.DataBarRule.class,
      name = "DATA_BAR_RULE"),
  @JsonSubTypes.Type(
      value = ConditionalFormattingRuleReport.IconSetRule.class,
      name = "ICON_SET_RULE"),
  @JsonSubTypes.Type(
      value = ConditionalFormattingRuleReport.UnsupportedRule.class,
      name = "UNSUPPORTED_RULE")
})
public sealed interface ConditionalFormattingRuleReport
    permits ConditionalFormattingRuleReport.FormulaRule,
        ConditionalFormattingRuleReport.CellValueRule,
        ConditionalFormattingRuleReport.ColorScaleRule,
        ConditionalFormattingRuleReport.DataBarRule,
        ConditionalFormattingRuleReport.IconSetRule,
        ConditionalFormattingRuleReport.UnsupportedRule {

  /** Persisted priority value loaded from the workbook. */
  int priority();

  /** Persisted stop-if-true flag loaded from the workbook. */
  boolean stopIfTrue();

  /** Formula-driven conditional-formatting rule with one differential-style payload. */
  record FormulaRule(
      int priority, boolean stopIfTrue, String formula, DifferentialStyleReport style)
      implements ConditionalFormattingRuleReport {
    public FormulaRule {
      requirePriority(priority);
      Objects.requireNonNull(formula, "formula must not be null");
      if (formula.isBlank()) {
        throw new IllegalArgumentException("formula must not be blank");
      }
    }
  }

  /** Cell-value comparison rule with one or two operands and one differential-style payload. */
  record CellValueRule(
      int priority,
      boolean stopIfTrue,
      ExcelComparisonOperator operator,
      String formula1,
      String formula2,
      DifferentialStyleReport style)
      implements ConditionalFormattingRuleReport {
    public CellValueRule {
      requirePriority(priority);
      Objects.requireNonNull(operator, "operator must not be null");
      Objects.requireNonNull(formula1, "formula1 must not be null");
      if (formula1.isBlank()) {
        throw new IllegalArgumentException("formula1 must not be blank");
      }
      if (formula2 != null && formula2.isBlank()) {
        throw new IllegalArgumentException("formula2 must not be blank");
      }
    }
  }

  /** Color-scale rule reported with thresholds and RGB color control points. */
  record ColorScaleRule(
      int priority,
      boolean stopIfTrue,
      List<ConditionalFormattingThresholdReport> thresholds,
      List<String> colors)
      implements ConditionalFormattingRuleReport {
    public ColorScaleRule {
      requirePriority(priority);
      thresholds = copyThresholds(thresholds);
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
      ConditionalFormattingThresholdReport minThreshold,
      ConditionalFormattingThresholdReport maxThreshold)
      implements ConditionalFormattingRuleReport {
    public DataBarRule {
      requirePriority(priority);
      color = ProtocolRgbColorSupport.normalizeRgbHex(color, "color");
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
      List<ConditionalFormattingThresholdReport> thresholds)
      implements ConditionalFormattingRuleReport {
    public IconSetRule {
      requirePriority(priority);
      Objects.requireNonNull(iconSet, "iconSet must not be null");
      thresholds = copyThresholds(thresholds);
    }
  }

  /** Loaded rule family that GridGrind can detect but not yet model directly. */
  record UnsupportedRule(int priority, boolean stopIfTrue, String kind, String detail)
      implements ConditionalFormattingRuleReport {
    public UnsupportedRule {
      requirePriority(priority);
      Objects.requireNonNull(kind, "kind must not be null");
      Objects.requireNonNull(detail, "detail must not be null");
      if (kind.isBlank()) {
        throw new IllegalArgumentException("kind must not be blank");
      }
      if (detail.isBlank()) {
        throw new IllegalArgumentException("detail must not be blank");
      }
    }
  }

  private static void requirePriority(int priority) {
    if (priority < 0) {
      throw new IllegalArgumentException("priority must not be negative");
    }
  }

  private static List<ConditionalFormattingThresholdReport> copyThresholds(
      List<ConditionalFormattingThresholdReport> thresholds) {
    Objects.requireNonNull(thresholds, "thresholds must not be null");
    List<ConditionalFormattingThresholdReport> copy = List.copyOf(thresholds);
    for (ConditionalFormattingThresholdReport threshold : copy) {
      Objects.requireNonNull(threshold, "thresholds must not contain null values");
    }
    return copy;
  }

  private static List<String> copyColors(List<String> colors) {
    Objects.requireNonNull(colors, "colors must not be null");
    List<String> copy = List.copyOf(colors);
    for (String color : copy) {
      ProtocolRgbColorSupport.normalizeRgbHex(color, "colors");
    }
    return copy;
  }
}
