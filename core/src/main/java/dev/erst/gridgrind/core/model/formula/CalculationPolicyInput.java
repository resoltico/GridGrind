package dev.erst.gridgrind.core.model.formula;

import com.fasterxml.jackson.annotation.JsonCreator;
import com.fasterxml.jackson.annotation.JsonIgnore;
import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.annotation.JsonProperty;
import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.metadata.ProtocolBooleanDefault;
import dev.erst.gridgrind.core.metadata.ProtocolField;
import java.util.Objects;

/**
 * Explicit request-side calculation policy covering evaluation, cache handling, and open-time
 * recalc.
 */
@CatalogType(
    id = "CalculationPolicyInput",
    summary =
        "Explicit formula-calculation policy. Use CalculationPolicyInput.defaults() for DO_NOT_CALCULATE with markRecalculateOnOpen=false. Use DEFERRED_CALCULATION to inspect capabilities without immediate evaluation, Use EVALUATE_ALL or EVALUATE_TARGETS for immediate server-side evaluation, CLEAR_CACHES_ONLY to strip persisted caches, or markRecalculateOnOpen=true when Excel-compatible clients should recalculate later.",
    displayOrder = 10)
public record CalculationPolicyInput(
    @ProtocolField(optional = true)
        @JsonInclude(
            value = JsonInclude.Include.CUSTOM,
            valueFilter = CalculationStrategyInput.DefaultFilter.class)
        CalculationStrategyInput strategy,
    @ProtocolField(optional = true, booleanDefault = ProtocolBooleanDefault.FALSE)
        @JsonInclude(JsonInclude.Include.NON_DEFAULT)
        boolean markRecalculateOnOpen) {
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

  /** Reads one calculation-policy block while applying the documented omission defaults. */
  @JsonCreator
  static CalculationPolicyInput create(
      @JsonProperty("strategy") CalculationStrategyInput strategy,
      @JsonProperty("markRecalculateOnOpen") Boolean markRecalculateOnOpen) {
    return new CalculationPolicyInput(
        strategy == null ? new CalculationStrategyInput.DoNotCalculate() : strategy,
        ProtocolBooleanDefault.FALSE.resolve(markRecalculateOnOpen));
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

  /** Returns whether this policy permits the constrained EVENT_READ execution family. */
  @JsonIgnore
  public boolean allowsEventRead() {
    return strategy instanceof CalculationStrategyInput.DoNotCalculate && !markRecalculateOnOpen;
  }

  /** Returns whether this policy permits the constrained STREAMING_WRITE execution family. */
  @JsonIgnore
  public boolean allowsStreamingWrite() {
    return strategy instanceof CalculationStrategyInput.DoNotCalculate;
  }

  /** Returns whether this policy needs all mutation steps before observation steps. */
  @JsonIgnore
  public boolean requiresMutationPrefix() {
    return !(strategy instanceof CalculationStrategyInput.DoNotCalculate)
        && !(strategy instanceof CalculationStrategyInput.DeferredCalculation);
  }

  /** Custom Jackson inclusion filter that omits the standard do-not-calculate policy. */
  public static final class DefaultFilter {
    @Override
    public boolean equals(Object other) {
      return other == null
          || (other instanceof CalculationPolicyInput calculation && calculation.isDefault());
    }

    @Override
    public int hashCode() {
      return DefaultFilter.class.hashCode();
    }
  }
}
