package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.excel.foundation.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.foundation.ExcelConditionalFormattingIconSet;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.Optional;

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
  record FormulaRule(
      String formula,
      boolean stopIfTrue,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<DifferentialStyleInput> style)
      implements ConditionalFormattingRuleInput {
    public FormulaRule {
      Objects.requireNonNull(formula, "formula must not be null");
      Objects.requireNonNull(style, "style must not be null");
      if (formula.isBlank()) {
        throw new IllegalArgumentException("formula must not be blank");
      }
      requireStyleOrStopBarrier(style, stopIfTrue);
      FormulaInputSecurity.rejectDde(formula); // LIM-027
    }
  }

  /** Cell-value comparison rule with one or two operands and one differential-style payload. */
  record CellValueRule(
      ExcelComparisonOperator operator,
      String formula1,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<String> formula2,
      boolean stopIfTrue,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<DifferentialStyleInput> style)
      implements ConditionalFormattingRuleInput {
    public CellValueRule {
      Objects.requireNonNull(operator, "operator must not be null");
      Objects.requireNonNull(formula1, "formula1 must not be null");
      Objects.requireNonNull(formula2, "formula2 must not be null");
      Objects.requireNonNull(style, "style must not be null");
      if (formula1.isBlank()) {
        throw new IllegalArgumentException("formula1 must not be blank");
      }
      requireStyleOrStopBarrier(style, stopIfTrue);
      FormulaInputSecurity.rejectDde(formula1); // LIM-027
      formula2 = normalizeOptionalComparisonUpperBound(operator, formula2);
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
      if (widthMin < ProtocolConstraintValues.DATA_BAR_WIDTH_MIN
          || widthMin > ProtocolConstraintValues.DATA_BAR_WIDTH_MAX) {
        throw new IllegalArgumentException("widthMin must be between 0 and 100 inclusive");
      }
      if (widthMax < ProtocolConstraintValues.DATA_BAR_WIDTH_MIN
          || widthMax > ProtocolConstraintValues.DATA_BAR_WIDTH_MAX) {
        throw new IllegalArgumentException("widthMax must be between 0 and 100 inclusive");
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
      boolean stopIfTrue,
      int rank,
      boolean percent,
      boolean bottom,
      @JsonInclude(JsonInclude.Include.NON_ABSENT) Optional<DifferentialStyleInput> style)
      implements ConditionalFormattingRuleInput {
    public Top10Rule {
      Objects.requireNonNull(style, "style must not be null");
      if (rank <= 0) {
        throw new IllegalArgumentException("rank must be greater than 0");
      }
      requireStyleOrStopBarrier(style, stopIfTrue);
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

  private static Optional<String> normalizeOptionalComparisonUpperBound(
      ExcelComparisonOperator operator, Optional<String> formula2) {
    if (operator == ExcelComparisonOperator.BETWEEN
        || operator == ExcelComparisonOperator.NOT_BETWEEN) {
      String upperBound = Objects.requireNonNullElse(formula2.orElse(null), "").trim();
      if (upperBound.isBlank()) {
        throw new IllegalArgumentException(
            "formula2 must not be blank for " + operator.name().toLowerCase(Locale.ROOT));
      }
      FormulaInputSecurity.rejectDde(upperBound); // LIM-027
      return Optional.of(upperBound);
    }
    if (formula2.isPresent()) {
      throw new IllegalArgumentException(
          "formula2 must be omitted unless operator is BETWEEN or NOT_BETWEEN");
    }
    return Optional.empty();
  }

  private static void requireStyleOrStopBarrier(
      Optional<DifferentialStyleInput> style, boolean stopIfTrue) {
    if (style.isEmpty() && !stopIfTrue) {
      throw new IllegalArgumentException(
          "style is required unless stopIfTrue creates a conditional-formatting barrier");
    }
  }
}
