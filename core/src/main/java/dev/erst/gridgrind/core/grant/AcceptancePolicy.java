package dev.erst.gridgrind.core.grant;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import dev.erst.gridgrind.core.plan.AssertionStep;
import dev.erst.gridgrind.core.plan.InspectionStep;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Objects;
import java.util.Set;

/** Host-required acceptance policy that authored plans cannot weaken. */
@Union(discriminator = "type")
public sealed interface AcceptancePolicy
    permits AcceptancePolicy.MinimumOnly, AcceptancePolicy.RequireAll {
  /** Requires the non-disableable staged-artifact structural verification minimum. */
  @Variant(
      id = "MINIMUM_ONLY",
      summary = "Require the non-disableable staged-artifact structural minimum.")
  record MinimumOnly() implements AcceptancePolicy {}

  /** Combines independently typed host requirements without nullable policy fields. */
  @Variant(id = "REQUIRE_ALL", summary = "Enforce every declared acceptance requirement.")
  record RequireAll(List<Requirement> requirements) implements AcceptancePolicy {
    public RequireAll {
      Objects.requireNonNull(requirements, "requirements must not be null");
      if (requirements.isEmpty()) {
        throw new IllegalArgumentException("requirements must not be empty; use minimum() instead");
      }
      requirements = List.copyOf(requirements);
    }
  }

  /** One independently enforced host acceptance requirement. */
  @Union(discriminator = "type")
  sealed interface Requirement
      permits Requirement.TerminalAssertions,
          Requirement.PreserveInspectionFacts,
          Requirement.PreserveOpaqueOoxmlParts,
          Requirement.RequireCalculation {
    /** Requires these host-owned assertions after plan execution and before publication. */
    @Variant(
        id = "TERMINAL_ASSERTIONS",
        summary = "Require these assertions before final publication.")
    record TerminalAssertions(List<AssertionStep> assertions) implements Requirement {
      public TerminalAssertions {
        Objects.requireNonNull(assertions, "assertions must not be null");
        if (assertions.isEmpty()) {
          throw new IllegalArgumentException("assertions must not be empty; use minimum() instead");
        }
        Set<dev.erst.gridgrind.core.vocabulary.StepId> stepIds = new LinkedHashSet<>();
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
    @Variant(
        id = "PRESERVE_INSPECTION_FACTS",
        summary = "Preserve the selected pre-mutation inspection facts.")
    record PreserveInspectionFacts(List<InspectionStep> inspections) implements Requirement {
      public PreserveInspectionFacts {
        Objects.requireNonNull(inspections, "inspections must not be null");
        if (inspections.isEmpty()) {
          throw new IllegalArgumentException("inspections must not be empty");
        }
        Set<dev.erst.gridgrind.core.vocabulary.StepId> stepIds = new LinkedHashSet<>();
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
    @Variant(
        id = "PRESERVE_OPAQUE_OOXML_PARTS",
        summary = "Preserve the named opaque OOXML parts byte-for-byte.")
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
    @Variant(
        id = "REQUIRE_CALCULATION",
        summary = "Require strict all-formula evaluation before publication.")
    record RequireCalculation() implements Requirement {}
  }

  /** Returns the product minimum acceptance policy. */
  static AcceptancePolicy minimum() {
    return new MinimumOnly();
  }
}
