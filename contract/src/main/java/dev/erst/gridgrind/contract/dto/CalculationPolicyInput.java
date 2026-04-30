package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonIgnore;
import com.fasterxml.jackson.annotation.JsonProperty;
import java.util.Objects;

/**
 * Explicit request-side calculation policy covering evaluation, cache handling, and open-time
 * recalc.
 */
public record CalculationPolicyInput(
    CalculationStrategyInput strategy, boolean markRecalculateOnOpen) {
  /** Returns the default do-not-calculate policy with no open-time recalculation request. */
  public static CalculationPolicyInput defaults() {
    return new CalculationPolicyInput(new CalculationStrategyInput.DoNotCalculate(), false);
  }

  /** Returns a calculation policy that customizes only the strategy. */
  public static CalculationPolicyInput strategy(CalculationStrategyInput strategy) {
    return new CalculationPolicyInput(strategy, false);
  }

  public CalculationPolicyInput {
    Objects.requireNonNull(strategy, "strategy must not be null");
  }

  /** Reads one calculation policy while applying the documented request defaults. */
  @JsonCreator
  public CalculationPolicyInput(
      @JsonProperty("strategy") CalculationStrategyInput strategy,
      @JsonProperty("markRecalculateOnOpen") Boolean markRecalculateOnOpen) {
    this(
        strategy,
        Objects.requireNonNull(markRecalculateOnOpen, "markRecalculateOnOpen must not be null")
            .booleanValue());
  }

  /** Returns whether this policy normalizes to the default do-not-calculate behavior. */
  @JsonIgnore
  public boolean isDefault() {
    return strategy instanceof CalculationStrategyInput.DoNotCalculate && !markRecalculateOnOpen;
  }

  /** Returns the normalized strategy after applying GridGrind defaults. */
  @JsonIgnore
  public CalculationStrategyInput effectiveStrategy() {
    return strategy;
  }
}
