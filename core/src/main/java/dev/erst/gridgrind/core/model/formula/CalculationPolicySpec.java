package dev.erst.gridgrind.core.model.formula;

import dev.erst.gridgrind.core.metadata.CatalogType;
import dev.erst.gridgrind.core.metadata.ProtocolField;
import dev.erst.gridgrind.core.metadata.WireDefault;
import java.util.Objects;

/**
 * Explicit request-side calculation policy covering evaluation, cache handling, and open-time
 * recalc.
 */
@CatalogType(
    id = "CalculationPolicySpec",
    summary =
        "Explicit formula-calculation policy. Use CalculationPolicySpec.defaults() for DO_NOT_CALCULATE with markRecalculateOnOpen=false. Use DEFERRED_CALCULATION to inspect capabilities without immediate evaluation, Use EVALUATE_ALL or EVALUATE_TARGETS for immediate server-side evaluation, CLEAR_CACHES_ONLY to strip persisted caches, or markRecalculateOnOpen=true when Excel-compatible clients should recalculate later.",
    displayOrder = 10)
public record CalculationPolicySpec(
    @ProtocolField(optional = true)
        @WireDefault(factory = CalculationStrategySpec.DoNotCalculate.class, omit = true)
        CalculationStrategySpec strategy,
    @ProtocolField(optional = true) @WireDefault(value = "false", omit = true)
        boolean markRecalculateOnOpen) {
  /** Returns the default do-not-calculate policy with no open-time recalculation request. */
  public static CalculationPolicySpec defaults() {
    return new CalculationPolicySpec(new CalculationStrategySpec.DoNotCalculate(), false);
  }

  /** Returns a calculation policy that customizes only the strategy. */
  public static CalculationPolicySpec strategy(CalculationStrategySpec strategy) {
    return new CalculationPolicySpec(strategy, false);
  }

  public CalculationPolicySpec {
    Objects.requireNonNull(strategy, "strategy must not be null");
  }

  /** Returns whether this policy normalizes to the default do-not-calculate behavior. */
  public boolean isDefault() {
    return strategy instanceof CalculationStrategySpec.DoNotCalculate && !markRecalculateOnOpen;
  }

  /** Returns the normalized strategy after applying GridGrind defaults. */
  public CalculationStrategySpec effectiveStrategy() {
    return strategy;
  }

  /** Returns whether this policy permits the constrained EVENT_READ execution family. */
  public boolean allowsEventRead() {
    return strategy instanceof CalculationStrategySpec.DoNotCalculate && !markRecalculateOnOpen;
  }

  /** Returns whether this policy permits the constrained STREAMING_WRITE execution family. */
  public boolean allowsStreamingWrite() {
    return strategy instanceof CalculationStrategySpec.DoNotCalculate;
  }

  /** Returns whether this policy needs all mutation steps before observation steps. */
  public boolean requiresMutationPrefix() {
    return !(strategy instanceof CalculationStrategySpec.DoNotCalculate)
        && !(strategy instanceof CalculationStrategySpec.DeferredCalculation);
  }
}
