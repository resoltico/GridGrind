package dev.erst.gridgrind.core.model.validation;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.model.style.DifferentialStyleReport;
import dev.erst.gridgrind.core.vocabulary.ExcelComparisonOperator;
import dev.erst.gridgrind.core.vocabulary.ExcelConditionalFormattingIconSet;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Protocol-facing factual report for one conditional-formatting rule loaded from a workbook. */
@Union(discriminator = "type")
public sealed interface ConditionalFormattingRuleReport
    permits ConditionalFormattingRuleReport.FormulaRule,
        ConditionalFormattingRuleReport.CellValueRule,
        ConditionalFormattingRuleReport.ColorScaleRule,
        ConditionalFormattingRuleReport.DataBarRule,
        ConditionalFormattingRuleReport.IconSetRule,
        ConditionalFormattingRuleReport.Top10Rule,
        ConditionalFormattingRuleReport.UnsupportedRule {

  /** Persisted priority value loaded from the workbook. */
  int priority();

  /** Persisted stop-if-true flag loaded from the workbook. */
  boolean stopIfTrue();

  /** Formula-driven conditional-formatting rule with one differential-style payload. */
  @Variant(
      id = "FORMULA_RULE",
      summary = "Formula-driven conditional-formatting rule with one differential-style payload.")
  record FormulaRule(
      int priority, boolean stopIfTrue, String formula, Optional<DifferentialStyleReport> style)
      implements ConditionalFormattingRuleReport {
    public FormulaRule {
      requirePriority(priority);
      Objects.requireNonNull(formula, "formula must not be null");
      Objects.requireNonNull(style, "style must not be null");
      TextValidation.requireNonBlank(formula, "formula");
    }
  }

  /** Cell-value comparison rule with one or two operands and one differential-style payload. */
  @Variant(
      id = "CELL_VALUE_RULE",
      summary =
          "Cell-value comparison rule with one or two operands and one differential-style payload.")
  record CellValueRule(
      int priority,
      boolean stopIfTrue,
      ExcelComparisonOperator operator,
      String formula1,
      Optional<String> formula2,
      Optional<DifferentialStyleReport> style)
      implements ConditionalFormattingRuleReport {
    public CellValueRule {
      requirePriority(priority);
      Objects.requireNonNull(operator, "operator must not be null");
      Objects.requireNonNull(formula1, "formula1 must not be null");
      Objects.requireNonNull(formula2, "formula2 must not be null");
      Objects.requireNonNull(style, "style must not be null");
      TextValidation.requireNonBlank(formula1, "formula1");
      if (formula2.isPresent() && formula2.orElseThrow().isBlank()) {
        throw new IllegalArgumentException("formula2 must not be blank when provided");
      }
    }
  }

  /** Color-scale rule reported with thresholds and RGB color control points. */
  @Variant(
      id = "COLOR_SCALE_RULE",
      summary = "Color-scale rule reported with thresholds and RGB color control points.")
  record ColorScaleRule(
      int priority,
      boolean stopIfTrue,
      List<ConditionalFormattingThresholdReport> thresholds,
      List<RgbHex> colors)
      implements ConditionalFormattingRuleReport {
    public ColorScaleRule {
      requirePriority(priority);
      thresholds = copyThresholds(thresholds);
      Objects.requireNonNull(colors, "colors must not be null");
      colors = List.copyOf(colors);
    }
  }

  /** Data-bar rule reported with thresholds, widths, and RGB fill color. */
  @Variant(
      id = "DATA_BAR_RULE",
      summary = "Data-bar rule reported with thresholds, widths, and RGB fill color.")
  record DataBarRule(
      int priority,
      boolean stopIfTrue,
      Optional<RgbHex> color,
      boolean iconOnly,
      int widthMin,
      int widthMax,
      ConditionalFormattingThresholdReport minThreshold,
      ConditionalFormattingThresholdReport maxThreshold)
      implements ConditionalFormattingRuleReport {
    public DataBarRule {
      requirePriority(priority);
      Objects.requireNonNull(color, "color must not be null");
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
  @Variant(
      id = "ICON_SET_RULE",
      summary = "Icon-set rule reported with persisted icon set, direction flags, and thresholds.")
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

  /** Top-10 rule reported with rank, percent, and bottom-state flags plus differential style. */
  @Variant(
      id = "TOP10_RULE",
      summary =
          "Top-10 rule reported with rank, percent, and bottom-state flags plus differential style.")
  record Top10Rule(
      int priority,
      boolean stopIfTrue,
      int rank,
      boolean percent,
      boolean bottom,
      Optional<DifferentialStyleReport> style)
      implements ConditionalFormattingRuleReport {
    public Top10Rule {
      requirePriority(priority);
      Objects.requireNonNull(style, "style must not be null");
      if (rank < 0) {
        throw new IllegalArgumentException("rank must not be negative");
      }
    }
  }

  /** Loaded rule family that GridGrind can detect but not yet model directly. */
  @Variant(
      id = "UNSUPPORTED_RULE",
      summary = "Loaded rule family that GridGrind can detect but not yet model directly.")
  record UnsupportedRule(int priority, boolean stopIfTrue, String kind, String detail)
      implements ConditionalFormattingRuleReport {
    public UnsupportedRule {
      requirePriority(priority);
      Objects.requireNonNull(kind, "kind must not be null");
      Objects.requireNonNull(detail, "detail must not be null");
      TextValidation.requireNonBlank(kind, "kind");
      TextValidation.requireNonBlank(detail, "detail");
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
}
