package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.engine.api.GridGrindExecutionGrant;
import dev.erst.gridgrind.engine.api.GridGrindHostAcceptancePolicy;
import java.util.List;
import java.util.Objects;
import org.jspecify.annotations.Nullable;

/** Validates that the authored execution mode can meet every host-owned acceptance requirement. */
final class HostAcceptancePolicyValidator {
  private HostAcceptancePolicyValidator() {}

  static List<HostAcceptancePolicyViolationException> violations(
      WorkbookPlan request, GridGrindExecutionGrant executionGrant) {
    Objects.requireNonNull(request, "request must not be null");
    GridGrindExecutionGrant.Bounded grant =
        (GridGrindExecutionGrant.Bounded)
            Objects.requireNonNull(executionGrant, "executionGrant must not be null");
    return switch (grant.hostAcceptancePolicy()) {
      case GridGrindHostAcceptancePolicy.MinimumOnly _ -> List.of();
      case GridGrindHostAcceptancePolicy.RequireAll requireAll ->
          violationsFor(requireAll, request);
    };
  }

  private static List<HostAcceptancePolicyViolationException> violationsFor(
      GridGrindHostAcceptancePolicy.RequireAll requireAll, WorkbookPlan request) {
    for (GridGrindHostAcceptancePolicy.Requirement requirement : requireAll.requirements()) {
      String message =
          switch (requirement) {
            case GridGrindHostAcceptancePolicy.Requirement.PreserveInspectionFacts _ ->
                fullXssfViolation(
                    request, "host semantic preservation requires execution.mode.type=FULL_XSSF");
            case GridGrindHostAcceptancePolicy.Requirement.PreserveOpaqueOoxmlParts _ ->
                opaquePartPreservationViolation(request);
            case GridGrindHostAcceptancePolicy.Requirement.RequireCalculation _ ->
                fullXssfViolation(
                    request, "host calculation requires execution.mode.type=FULL_XSSF");
            case GridGrindHostAcceptancePolicy.Requirement.TerminalAssertions _ -> null;
          };
      if (message != null) {
        return List.of(new HostAcceptancePolicyViolationException(message));
      }
    }
    return List.of();
  }

  private static @Nullable String opaquePartPreservationViolation(WorkbookPlan request) {
    String fullXssfViolation =
        fullXssfViolation(
            request, "host opaque-part preservation requires execution.mode.type=FULL_XSSF");
    if (fullXssfViolation != null) {
      return fullXssfViolation;
    }
    if (!(request.source() instanceof WorkbookPlan.WorkbookSource.ExistingFile existingFile)) {
      return "host opaque-part preservation requires an unencrypted EXISTING source workbook";
    }
    if (existingFile.security().isPresent()) {
      return "host opaque-part preservation does not support encrypted source workbooks";
    }
    if (request.persistence() instanceof WorkbookPlan.WorkbookPersistence.None) {
      return "host opaque-part preservation requires a persisted workbook artifact";
    }
    return null;
  }

  private static @Nullable String fullXssfViolation(WorkbookPlan request, String message) {
    return request.execution().mode() instanceof ExecutionModeInput.FullXssf ? null : message;
  }
}
