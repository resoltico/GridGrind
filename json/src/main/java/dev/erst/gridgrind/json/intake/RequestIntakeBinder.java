package dev.erst.gridgrind.json.intake;

import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.RequestBindingFailure;
import dev.erst.gridgrind.core.problem.RequestBoundFragments;
import dev.erst.gridgrind.core.problem.RequestIntake;
import dev.erst.gridgrind.core.problem.RequestStructuralProblem;
import dev.erst.gridgrind.core.problem.RequestWarning;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Optional;

/** Completes independently bound fragments and releases parser state into a pure intake result. */
final class RequestIntakeBinder {
  private RequestIntakeBinder() {}

  static RequestIntake bind(
      RequestBoundFragments boundFragments,
      List<RequestStructuralProblem> structuralProblems,
      List<RequestBindingFailure> bindingFailures,
      Optional<RequestJsonNode> rawRoot,
      List<RequestWarning> warnings) {
    Objects.requireNonNull(boundFragments, "boundFragments must not be null");
    List<RequestStructuralProblem> orderedProblems =
        RequestStructuralProblemOrder.order(List.copyOf(structuralProblems));
    Optional<RequestJsonNode> parsedRoot = Objects.requireNonNullElseGet(rawRoot, Optional::empty);
    List<RequestBindingFailure> discoveredFailures = new ArrayList<>(bindingFailures);
    Optional<WorkbookPlan> completePlan = Optional.empty();
    if (orderedProblems.isEmpty() && discoveredFailures.isEmpty()) {
      try {
        completePlan = boundFragments.completePlan();
      } catch (IllegalArgumentException exception) {
        discoveredFailures.add(RequestConstructorValidation.problem(exception, "", false));
      }
    }
    return new RequestIntake(
        boundFragments,
        orderedProblems,
        discoveredFailures.stream()
            .map(failure -> locateBindingFailure(failure, parsedRoot))
            .toList(),
        completePlan,
        parsedRoot.map(RequestJsonTokenLocationSupport::index).orElseGet(Map::of),
        warnings);
  }

  private static RequestBindingFailure locateBindingFailure(
      RequestBindingFailure failure, Optional<RequestJsonNode> rawRoot) {
    Optional<Long> byteOffset =
        rawRoot.flatMap(
            root -> RequestJsonTokenLocationSupport.byteOffsetAt(root, failure.jsonPath()));
    return byteOffset.isPresent() ? failure.locatedAt(byteOffset) : failure;
  }
}
