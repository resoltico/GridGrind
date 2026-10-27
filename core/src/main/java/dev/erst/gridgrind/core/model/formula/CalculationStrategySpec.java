package dev.erst.gridgrind.core.model.formula;

import dev.erst.gridgrind.core.catalog.GridGrindProtocolTypeNames;
import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;

/** Explicit request-side strategy for formula calculation handling. */
@Union(discriminator = "type", catalog = true, displayOrder = 71)
public sealed interface CalculationStrategySpec {
  /** Stable SCREAMING_SNAKE_CASE discriminator used in the public contract. */
  String strategyType();

  /** Returns whether this strategy is the default no-calculation choice. */
  default boolean isDefault() {
    return this instanceof DoNotCalculate;
  }

  /** Do not run server-side formula evaluation or cache clearing. */
  @Variant(id = "DO_NOT_CALCULATE", summary = "Skip immediate server-side formula calculation.")
  record DoNotCalculate() implements CalculationStrategySpec {
    @Override
    public String strategyType() {
      return GridGrindProtocolTypeNames.calculationStrategyTypeName(DoNotCalculate.class);
    }
  }

  /** Inspect formula capabilities but defer all calculation to an Excel-compatible client. */
  @Variant(
      id = "DEFERRED_CALCULATION",
      summary =
          "Inspect formula capabilities and defer all calculation to an Excel-compatible client.")
  record DeferredCalculation() implements CalculationStrategySpec {
    @Override
    public String strategyType() {
      return GridGrindProtocolTypeNames.calculationStrategyTypeName(DeferredCalculation.class);
    }
  }

  /** Evaluate every formula cell reachable in the workbook after mutation steps complete. */
  @Variant(
      id = "EVALUATE_ALL",
      summary = "Preflight and evaluate every formula cell after mutation steps complete.")
  record EvaluateAll() implements CalculationStrategySpec {
    @Override
    public String strategyType() {
      return GridGrindProtocolTypeNames.calculationStrategyTypeName(EvaluateAll.class);
    }
  }

  /** Evaluate one explicit set of formula-cell targets after mutation steps complete. */
  @Variant(
      id = "EVALUATE_TARGETS",
      summary = "Preflight and evaluate the explicit formula-cell target list only.")
  record EvaluateTargets(List<FormulaCellTarget> cells) implements CalculationStrategySpec {
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
  @Variant(
      id = "REQUIRE_EVALUATION",
      summary =
          "Preflight and evaluate every formula cell, failing when any formula cannot be evaluated immediately.")
  record RequireEvaluation() implements CalculationStrategySpec {
    @Override
    public String strategyType() {
      return GridGrindProtocolTypeNames.calculationStrategyTypeName(RequireEvaluation.class);
    }
  }

  /** Clear persisted formula caches without attempting immediate evaluation. */
  @Variant(
      id = "CLEAR_CACHES_ONLY",
      summary = "Strip persisted formula caches without running immediate evaluation.")
  record ClearCachesOnly() implements CalculationStrategySpec {
    @Override
    public String strategyType() {
      return GridGrindProtocolTypeNames.calculationStrategyTypeName(ClearCachesOnly.class);
    }
  }
}
