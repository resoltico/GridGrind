package dev.erst.gridgrind.core.problem;

import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces.JsonLocation;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;

/** Bound request data and ordered findings, independent of parser state and JSON libraries. */
public record RequestIntake(
    RequestBoundFragments boundFragments,
    List<RequestStructuralProblem> structuralProblems,
    List<RequestBindingFailure> bindingFailures,
    Optional<WorkbookPlan> completePlan,
    Map<String, Long> tokenOffsets,
    List<RequestWarning> warnings) {
  public RequestIntake {
    Objects.requireNonNull(boundFragments, "boundFragments must not be null");
    structuralProblems = List.copyOf(structuralProblems);
    bindingFailures = List.copyOf(bindingFailures);
    Objects.requireNonNull(completePlan, "completePlan must not be null");
    tokenOffsets = Map.copyOf(tokenOffsets);
    warnings = List.copyOf(warnings);
    if (completePlan.isPresent() && (!structuralProblems.isEmpty() || !bindingFailures.isEmpty())) {
      throw new IllegalArgumentException(
          "An invalid request intake cannot contain a complete plan");
    }
  }

  /** Returns the plan after the caller has handled every intake finding. */
  public WorkbookPlan requireCompletePlan() {
    if (!structuralProblems.isEmpty()) {
      throw new IllegalStateException(
          "A structurally invalid request cannot produce a complete plan");
    }
    if (!bindingFailures.isEmpty()) {
      throw bindingFailures.getFirst().exception();
    }
    return completePlan.orElseThrow(
        () ->
            new IllegalStateException(
                "A structurally valid request must bind every required request fragment"));
  }

  /** Returns whether tolerant parsing and structural collection found no request-shape defects. */
  public boolean isStructurallyValid() {
    return structuralProblems.isEmpty();
  }

  /** Returns whether every fragment and complete-plan invariant bound successfully. */
  public boolean isBindable() {
    return isStructurallyValid() && bindingFailures.isEmpty() && completePlan.isPresent();
  }

  /** Returns the authored token offset for one canonical request path, when known. */
  public Optional<Long> byteOffsetAt(String jsonPath) {
    return Optional.ofNullable(
        tokenOffsets.get(Objects.requireNonNull(jsonPath, "jsonPath must not be null")));
  }

  /** Returns the strongest available source location for one canonical request path. */
  public JsonLocation jsonLocationAt(String jsonPath) {
    return byteOffsetAt(jsonPath)
        .map(offset -> JsonLocation.pathAtByteOffset(jsonPath, offset))
        .orElseGet(() -> JsonLocation.pathOnly(jsonPath));
  }
}
