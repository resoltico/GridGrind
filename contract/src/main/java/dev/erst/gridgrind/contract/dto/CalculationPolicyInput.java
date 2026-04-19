package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonIgnore;
import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonProperty;
import java.util.Objects;

/**
 * Explicit request-side calculation policy covering evaluation, cache handling, and open-time
 * recalc.
 */
public record CalculationPolicyInput(
    @JsonInclude(JsonInclude.Include.NON_NULL) CalculationStrategyInput strategy,
    @JsonInclude(JsonInclude.Include.NON_DEFAULT) boolean markRecalculateOnOpen) {
  public CalculationPolicyInput {
    strategy =
        Objects.requireNonNullElseGet(strategy, CalculationStrategyInput.DoNotCalculate::new);
  }

  /** Creates a calculation policy with the provided strategy and no open-time recalc flag. */
  public CalculationPolicyInput(CalculationStrategyInput strategy) {
    this(strategy, false);
  }

  /** Deserializes request and response payloads while treating omitted flags as false. */
  @JsonCreator
  public static CalculationPolicyInput create(
      @JsonProperty("strategy") CalculationStrategyInput strategy,
      @JsonProperty("markRecalculateOnOpen") Boolean markRecalculateOnOpen) {
    return new CalculationPolicyInput(strategy, Boolean.TRUE.equals(markRecalculateOnOpen));
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
