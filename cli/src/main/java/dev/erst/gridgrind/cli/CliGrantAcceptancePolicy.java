package dev.erst.gridgrind.cli;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.plan.AssertionStep;
import dev.erst.gridgrind.core.plan.InspectionStep;
import dev.erst.gridgrind.engine.api.GridGrindHostAcceptancePolicy;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/**
 * CLI-owned host acceptance policy translated without becoming part of an authored workbook plan.
 */
@Union(discriminator = "type")
public sealed interface CliGrantAcceptancePolicy
    permits CliGrantAcceptancePolicy.MinimumOnly, CliGrantAcceptancePolicy.RequireAll {
  /** Converts this CLI-owned policy document into the engine API policy. */
  GridGrindHostAcceptancePolicy toEnginePolicy();

  /** Requires only the product's non-disableable staged-artifact minimum. */
  @Variant(
      id = "MINIMUM_ONLY",
      summary = "Requires only the product's non-disableable staged-artifact minimum.")
  record MinimumOnly() implements CliGrantAcceptancePolicy {
    @Override
    public GridGrindHostAcceptancePolicy toEnginePolicy() {
      return GridGrindHostAcceptancePolicy.minimum();
    }
  }

  /** Combines independently typed host requirements without nullable policy fields. */
  @Variant(
      id = "REQUIRE_ALL",
      summary = "Combines independently typed host requirements without nullable policy fields.")
  record RequireAll(List<Requirement> requirements) implements CliGrantAcceptancePolicy {
    public RequireAll {
      requirements = copyValues(requirements, "requirements");
    }

    @Override
    public GridGrindHostAcceptancePolicy toEnginePolicy() {
      return new GridGrindHostAcceptancePolicy.RequireAll(
          requirements.stream().map(Requirement::toEngineRequirement).toList());
    }
  }

  /** One independently enforced host acceptance requirement. */
  @Union(discriminator = "type")
  sealed interface Requirement
      permits Requirement.TerminalAssertions,
          Requirement.PreserveInspectionFacts,
          CliGrantOpaqueOoxmlPartsRequirement,
          CliGrantRequiredCalculation {
    /** Converts this CLI-owned requirement into the engine API requirement. */
    GridGrindHostAcceptancePolicy.Requirement toEngineRequirement();

    /** Requires host-owned assertions after plan execution and before publication. */
    @Variant(
        id = "TERMINAL_ASSERTIONS",
        summary = "Requires host-owned assertions after plan execution and before publication.")
    record TerminalAssertions(List<AssertionStep> assertions) implements Requirement {
      public TerminalAssertions {
        assertions = copyValues(assertions, "assertions");
      }

      @Override
      public GridGrindHostAcceptancePolicy.Requirement toEngineRequirement() {
        return new GridGrindHostAcceptancePolicy.Requirement.TerminalAssertions(assertions);
      }
    }

    /** Requires selected inspection facts to remain semantically unchanged. */
    @Variant(
        id = "PRESERVE_INSPECTION_FACTS",
        summary = "Requires selected inspection facts to remain semantically unchanged.")
    record PreserveInspectionFacts(List<InspectionStep> inspections) implements Requirement {
      public PreserveInspectionFacts {
        inspections = copyValues(inspections, "inspections");
      }

      @Override
      public GridGrindHostAcceptancePolicy.Requirement toEngineRequirement() {
        return new GridGrindHostAcceptancePolicy.Requirement.PreserveInspectionFacts(inspections);
      }
    }
  }

  /** Validates and copies one grant-document collection without retaining mutable caller state. */
  static <T> List<T> copyValues(List<T> values, String fieldName) {
    Objects.requireNonNull(values, fieldName + " must not be null");
    List<T> copied = new ArrayList<>(values.size());
    for (T value : values) {
      copied.add(Objects.requireNonNull(value, fieldName + " must not contain null values"));
    }
    return List.copyOf(copied);
  }
}
