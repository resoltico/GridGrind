package dev.erst.gridgrind.contract.dto;

import com.fasterxml.jackson.annotation.JsonIgnore;
import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeInfo;
import dev.erst.gridgrind.contract.catalog.GridGrindProtocolTypeNames;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;

/** Explicit request-side strategy for formula calculation handling. */
@JsonTypeInfo(use = JsonTypeInfo.Id.NAME, property = "type")
@JsonSubTypes({
  @JsonSubTypes.Type(
      value = CalculationStrategyInput.DoNotCalculate.class,
      name = "DO_NOT_CALCULATE"),
  @JsonSubTypes.Type(
      value = CalculationStrategyInput.DeferredCalculation.class,
      name = "DEFERRED_CALCULATION"),
  @JsonSubTypes.Type(value = CalculationStrategyInput.EvaluateAll.class, name = "EVALUATE_ALL"),
  @JsonSubTypes.Type(
      value = CalculationStrategyInput.EvaluateTargets.class,
      name = "EVALUATE_TARGETS"),
  @JsonSubTypes.Type(
      value = CalculationStrategyInput.RequireEvaluation.class,
      name = "REQUIRE_EVALUATION"),
  @JsonSubTypes.Type(
      value = CalculationStrategyInput.ClearCachesOnly.class,
      name = "CLEAR_CACHES_ONLY")
})
public sealed interface CalculationStrategyInput {
  /** Stable SCREAMING_SNAKE_CASE discriminator used in the public contract. */
  String strategyType();

  /** Returns whether this strategy is the default no-calculation choice. */
  @JsonIgnore
  default boolean isDefault() {
    return this instanceof DoNotCalculate;
  }

  /** Custom Jackson inclusion filter that omits the standard DO_NOT_CALCULATE strategy. */
  final class DefaultFilter {
    @Override
    public boolean equals(Object other) {
      return other == null
          || (other instanceof CalculationStrategyInput strategy && strategy.isDefault());
    }

    @Override
    public int hashCode() {
      return DefaultFilter.class.hashCode();
    }
  }

  /** Do not run server-side formula evaluation or cache clearing. */
  record DoNotCalculate() implements CalculationStrategyInput {
    @Override
    public String strategyType() {
      return GridGrindProtocolTypeNames.calculationStrategyTypeName(DoNotCalculate.class);
    }
  }

  /** Inspect formula capabilities but defer all calculation to an Excel-compatible client. */
  record DeferredCalculation() implements CalculationStrategyInput {
    @Override
    public String strategyType() {
      return GridGrindProtocolTypeNames.calculationStrategyTypeName(DeferredCalculation.class);
    }
  }

  /** Evaluate every formula cell reachable in the workbook after mutation steps complete. */
  record EvaluateAll() implements CalculationStrategyInput {
    @Override
    public String strategyType() {
      return GridGrindProtocolTypeNames.calculationStrategyTypeName(EvaluateAll.class);
    }
  }

  /** Evaluate one explicit set of formula-cell targets after mutation steps complete. */
  record EvaluateTargets(List<FormulaCellTarget> cells) implements CalculationStrategyInput {
    public EvaluateTargets {
      Objects.requireNonNull(cells, "cells must not be null");
      if (cells.isEmpty()) {
        throw new IllegalArgumentException("cells must not be empty");
      }
      Set<FormulaCellTarget> deduplicated = new LinkedHashSet<>();
      for (FormulaCellTarget cell : cells) {
        deduplicated.add(Objects.requireNonNull(cell, "cells must not contain nulls"));
      }
      if (deduplicated.size() != cells.size()) {
        throw new IllegalArgumentException("cells must not contain duplicates");
      }
      cells = List.copyOf(cells);
    }

    @Override
    public String strategyType() {
      return GridGrindProtocolTypeNames.calculationStrategyTypeName(EvaluateTargets.class);
    }
  }

  /** Evaluate every formula and fail when any formula cannot be evaluated immediately. */
  record RequireEvaluation() implements CalculationStrategyInput {
    @Override
    public String strategyType() {
      return GridGrindProtocolTypeNames.calculationStrategyTypeName(RequireEvaluation.class);
    }
  }

  /** Clear persisted formula caches without attempting immediate evaluation. */
  record ClearCachesOnly() implements CalculationStrategyInput {
    @Override
    public String strategyType() {
      return GridGrindProtocolTypeNames.calculationStrategyTypeName(ClearCachesOnly.class);
    }
  }
}
