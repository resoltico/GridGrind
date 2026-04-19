package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.excel.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingIconSet;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/** Protocol-facing authored conditional-formatting rule families. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(
      value = ConditionalFormattingRuleInput.FormulaRule.class,
      name = "FORMULA_RULE"),
  @JsonSubTypes.Type(
      value = ConditionalFormattingRuleInput.CellValueRule.class,
      name = "CELL_VALUE_RULE"),
  @JsonSubTypes.Type(
      value = ConditionalFormattingRuleInput.ColorScaleRule.class,
      name = "COLOR_SCALE_RULE"),
  @JsonSubTypes.Type(
      value = ConditionalFormattingRuleInput.DataBarRule.class,
      name = "DATA_BAR_RULE"),
  @JsonSubTypes.Type(
      value = ConditionalFormattingRuleInput.IconSetRule.class,
      name = "ICON_SET_RULE"),
  @JsonSubTypes.Type(value = ConditionalFormattingRuleInput.Top10Rule.class, name = "TOP10_RULE")
})
public sealed interface ConditionalFormattingRuleInput
    permits ConditionalFormattingRuleInput.FormulaRule,
        ConditionalFormattingRuleInput.CellValueRule,
        ConditionalFormattingRuleInput.ColorScaleRule,
        ConditionalFormattingRuleInput.DataBarRule,
        ConditionalFormattingRuleInput.IconSetRule,
        ConditionalFormattingRuleInput.Top10Rule {

  /** Formula-driven conditional-formatting rule with one differential-style payload. */
  record FormulaRule(String formula, boolean stopIfTrue, DifferentialStyleInput style)
      implements ConditionalFormattingRuleInput {
    public FormulaRule {
      Objects.requireNonNull(formula, "formula must not be null");
      if (formula.isBlank()) {
        throw new IllegalArgumentException("formula must not be blank");
      }
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
    }
  }

  /** Color-scale rule with ordered thresholds and control-point colors. */
  record ColorScaleRule(
      boolean stopIfTrue,
      List<ConditionalFormattingThresholdInput> thresholds,
      List<ColorInput> colors)
      implements ConditionalFormattingRuleInput {
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

  /** Data-bar rule with thresholds, widths, and a fill color. */
  record DataBarRule(
      boolean stopIfTrue,
      ColorInput color,
      boolean iconOnly,
      int widthMin,
      int widthMax,
      ConditionalFormattingThresholdInput minThreshold,
      ConditionalFormattingThresholdInput maxThreshold)
      implements ConditionalFormattingRuleInput {
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

  /** Icon-set rule with authored icon set and thresholds. */
  record IconSetRule(
      boolean stopIfTrue,
      ExcelConditionalFormattingIconSet iconSet,
      boolean iconOnly,
      boolean reversed,
      List<ConditionalFormattingThresholdInput> thresholds)
      implements ConditionalFormattingRuleInput {
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

  /** Top-N or bottom-N conditional-format rule with a differential style. */
  record Top10Rule(
      boolean stopIfTrue, int rank, boolean percent, boolean bottom, DifferentialStyleInput style)
      implements ConditionalFormattingRuleInput {
    public Top10Rule {
      if (rank <= 0) {
        throw new IllegalArgumentException("rank must be greater than 0");
      }
    }
  }

  private static List<ConditionalFormattingThresholdInput> copyThresholds(
      List<ConditionalFormattingThresholdInput> thresholds, String fieldName) {
    List<ConditionalFormattingThresholdInput> copy =
        List.copyOf(Objects.requireNonNull(thresholds, fieldName + " must not be null"));
    for (ConditionalFormattingThresholdInput threshold : copy) {
      Objects.requireNonNull(threshold, fieldName + " must not contain null values");
    }
    return copy;
  }

  private static List<ColorInput> copyColors(List<ColorInput> colors, String fieldName) {
    List<ColorInput> copy =
        new ArrayList<>(
            List.copyOf(Objects.requireNonNull(colors, fieldName + " must not be null")));
    for (ColorInput color : copy) {
      Objects.requireNonNull(color, fieldName + " must not contain null values");
    }
    return List.copyOf(copy);
  }
}
