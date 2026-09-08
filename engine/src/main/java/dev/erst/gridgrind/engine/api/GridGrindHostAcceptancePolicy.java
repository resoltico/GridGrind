package dev.erst.gridgrind.engine.api;

import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;

/** Host-required acceptance policy that authored plans cannot weaken. */
public sealed interface GridGrindHostAcceptancePolicy
    permits GridGrindHostAcceptancePolicy.MinimumOnly, GridGrindHostAcceptancePolicy.RequireAll {
  /** Requires the non-disableable staged-artifact structural verification minimum. */
  record MinimumOnly() implements GridGrindHostAcceptancePolicy {}

  /** Combines independently typed host requirements without nullable policy fields. */
  record RequireAll(List<Requirement> requirements) implements GridGrindHostAcceptancePolicy {
    public RequireAll {
      Objects.requireNonNull(requirements, "requirements must not be null");
      if (requirements.isEmpty()) {
        throw new IllegalArgumentException("requirements must not be empty; use minimum() instead");
      }
      requirements = List.copyOf(requirements);
    }
  }

  /** One independently enforced host acceptance requirement. */
  sealed interface Requirement
      permits Requirement.TerminalAssertions,
          Requirement.PreserveInspectionFacts,
          Requirement.PreserveOpaqueOoxmlParts,
          Requirement.RequireCalculation {
    /** Requires these host-owned assertions after plan execution and before publication. */
    record TerminalAssertions(List<AssertionStep> assertions) implements Requirement {
      public TerminalAssertions {
        Objects.requireNonNull(assertions, "assertions must not be null");
        if (assertions.isEmpty()) {
          throw new IllegalArgumentException("assertions must not be empty; use minimum() instead");
        }
        Set<String> stepIds = new LinkedHashSet<>();
        for (AssertionStep assertion : assertions) {
          AssertionStep required =
              Objects.requireNonNull(assertion, "assertions must not contain null values");
          if (!stepIds.add(required.stepId())) {
            throw new IllegalArgumentException("assertions must not reuse step ids");
          }
        }
        assertions = List.copyOf(assertions);
      }
    }

    /**
     * Requires selected facts captured before mutation to be semantically unchanged at acceptance.
     */
    record PreserveInspectionFacts(List<InspectionStep> inspections) implements Requirement {
      public PreserveInspectionFacts {
        Objects.requireNonNull(inspections, "inspections must not be null");
        if (inspections.isEmpty()) {
          throw new IllegalArgumentException("inspections must not be empty");
        }
        Set<String> stepIds = new LinkedHashSet<>();
        for (InspectionStep inspection : inspections) {
          InspectionStep required =
              Objects.requireNonNull(inspection, "inspections must not contain null values");
          if (!stepIds.add(required.stepId())) {
            throw new IllegalArgumentException("inspections must not reuse step ids");
          }
        }
        inspections = List.copyOf(inspections);
      }
    }

    /** Requires named opaque OOXML package parts to remain byte-identical after serialization. */
    record PreserveOpaqueOoxmlParts(List<String> partNames) implements Requirement {
      public PreserveOpaqueOoxmlParts {
        Objects.requireNonNull(partNames, "partNames must not be null");
        if (partNames.isEmpty()) {
          throw new IllegalArgumentException("partNames must not be empty");
        }
        Set<String> distinct = new LinkedHashSet<>();
        for (String partName : partNames) {
          String required =
              Objects.requireNonNull(partName, "partNames must not contain null values");
          if (!isOoxmlPartName(required)) {
            throw new IllegalArgumentException(
                "partNames must contain normalized OOXML part names");
          }
          if (!distinct.add(required)) {
            throw new IllegalArgumentException("partNames must not contain duplicates");
          }
        }
        partNames = List.copyOf(partNames);
      }

      private static boolean isOoxmlPartName(String value) {
        return value.startsWith("/")
            && value.length() > 1
            && !value.endsWith("/")
            && !value.contains("\\")
            && !value.contains("//")
            && value.chars().noneMatch(character -> character == '\u0000')
            && java.util.Arrays.stream(value.substring(1).split("/", -1))
                .noneMatch(segment -> ".".equals(segment) || "..".equals(segment));
      }
    }

    /**
     * Requires a strict all-formula evaluation after every plan mutation and before publication.
     */
    record RequireCalculation() implements Requirement {}
  }

  /** Returns the product minimum acceptance policy. */
  static GridGrindHostAcceptancePolicy minimum() {
    return new MinimumOnly();
  }
}
