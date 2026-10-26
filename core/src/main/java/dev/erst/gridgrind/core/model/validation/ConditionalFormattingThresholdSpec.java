package dev.erst.gridgrind.core.model.validation;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.model.formula.FormulaInputSecurity;

/** Tagged authoring threshold variants for advanced conditional-formatting rules. */
@Union(discriminator = "type", catalog = true, displayOrder = 65)
public sealed interface ConditionalFormattingThresholdSpec
    permits dev.erst.gridgrind.core.model.validation.ConditionalFormattingThresholdSpec.Min,
        dev.erst.gridgrind.core.model.validation.ConditionalFormattingThresholdSpec.Max,
        dev.erst.gridgrind.core.model.validation.ConditionalFormattingThresholdSpec.Numeric,
        dev.erst.gridgrind.core.model.validation.ConditionalFormattingThresholdSpec.Percent,
        dev.erst.gridgrind.core.model.validation.ConditionalFormattingThresholdSpec.Percentile,
        dev.erst.gridgrind.core.model.validation.ConditionalFormattingThresholdSpec.Formula {

  /** Uses the minimum value observed in the applied range. */
  @Variant(id = "MIN", summary = "Use the minimum value observed in the applied range.")
  record Min() implements ConditionalFormattingThresholdSpec {}

  /** Uses the maximum value observed in the applied range. */
  @Variant(id = "MAX", summary = "Use the maximum value observed in the applied range.")
  record Max() implements ConditionalFormattingThresholdSpec {}

  /** Uses one finite numeric threshold. */
  @Variant(id = "NUMBER", summary = "Use one finite numeric threshold.")
  record Numeric(double value) implements ConditionalFormattingThresholdSpec {
    public Numeric {
      requireFinite(value, "value");
    }
  }

  /** Uses one finite percentage threshold in the inclusive range {@code [0, 100]}. */
  @Variant(id = "PERCENT", summary = "Use one finite percentage threshold from 0 through 100.")
  record Percent(double value) implements ConditionalFormattingThresholdSpec {
    public Percent {
      requirePercentage(value, "value");
    }
  }

  /** Uses one finite percentile threshold in the inclusive range {@code [0, 100]}. */
  @Variant(id = "PERCENTILE", summary = "Use one finite percentile threshold from 0 through 100.")
  record Percentile(double value) implements ConditionalFormattingThresholdSpec {
    public Percentile {
      requirePercentage(value, "value");
    }
  }

  /** Uses one nonblank, DDE-safe formula threshold. */
  @Variant(id = "FORMULA", summary = "Use one nonblank, DDE-safe formula threshold.")
  record Formula(
      @dev.erst.gridgrind.core.metadata.FieldRule(
              kind = dev.erst.gridgrind.core.metadata.FieldRule.Kind.NON_BLANK)
          String formula)
      implements ConditionalFormattingThresholdSpec {
    public Formula {
      if (formula == null || formula.isBlank()) {
        throw new IllegalArgumentException("formula must not be blank");
      }
      FormulaInputSecurity.rejectDde(formula); // LIM-027
    }
  }

  private static void requireFinite(double value, String fieldName) {
    if (!Double.isFinite(value)) {
      throw new IllegalArgumentException(fieldName + " must be finite");
    }
  }

  private static void requirePercentage(double value, String fieldName) {
    requireFinite(value, fieldName);
    if (value < 0.0d || value > 100.0d) {
      throw new IllegalArgumentException(fieldName + " must be between 0 and 100 inclusive");
    }
  }
}
