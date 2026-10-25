package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;

/** Tagged authoring threshold variants for advanced conditional-formatting rules. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(value = ConditionalFormattingThresholdInput.Min.class, name = "MIN"),
  @JsonSubTypes.Type(value = ConditionalFormattingThresholdInput.Max.class, name = "MAX"),
  @JsonSubTypes.Type(value = ConditionalFormattingThresholdInput.Numeric.class, name = "NUMBER"),
  @JsonSubTypes.Type(value = ConditionalFormattingThresholdInput.Percent.class, name = "PERCENT"),
  @JsonSubTypes.Type(
      value = ConditionalFormattingThresholdInput.Percentile.class,
      name = "PERCENTILE"),
  @JsonSubTypes.Type(value = ConditionalFormattingThresholdInput.Formula.class, name = "FORMULA")
})
public sealed interface ConditionalFormattingThresholdInput
    permits ConditionalFormattingThresholdInput.Min,
        ConditionalFormattingThresholdInput.Max,
        ConditionalFormattingThresholdInput.Numeric,
        ConditionalFormattingThresholdInput.Percent,
        ConditionalFormattingThresholdInput.Percentile,
        ConditionalFormattingThresholdInput.Formula {

  /** Uses the minimum value observed in the applied range. */
  record Min() implements ConditionalFormattingThresholdInput {}

  /** Uses the maximum value observed in the applied range. */
  record Max() implements ConditionalFormattingThresholdInput {}

  /** Uses one finite numeric threshold. */
  record Numeric(double value) implements ConditionalFormattingThresholdInput {
    public Numeric {
      requireFinite(value, "value");
    }
  }

  /** Uses one finite percentage threshold in the inclusive range {@code [0, 100]}. */
  record Percent(double value) implements ConditionalFormattingThresholdInput {
    public Percent {
      requirePercentage(value, "value");
    }
  }

  /** Uses one finite percentile threshold in the inclusive range {@code [0, 100]}. */
  record Percentile(double value) implements ConditionalFormattingThresholdInput {
    public Percentile {
      requirePercentage(value, "value");
    }
  }

  /** Uses one nonblank, DDE-safe formula threshold. */
  record Formula(String formula) implements ConditionalFormattingThresholdInput {
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
