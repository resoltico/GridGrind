package dev.erst.gridgrind.contract.json;

import dev.erst.gridgrind.contract.dto.RequestWarning;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.core.problem.RequestBindingFailure;
import dev.erst.gridgrind.core.problem.RequestBoundFragments;
import dev.erst.gridgrind.core.problem.RequestIntake;
import dev.erst.gridgrind.core.problem.RequestStructuralProblem;
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
      RequestDiagnosticRedactor diagnosticRedactor,
      Optional<RequestJsonNode> rawRoot,
      List<RequestWarning> warnings) {
    Objects.requireNonNull(boundFragments, "boundFragments must not be null");
    List<RequestStructuralProblem> orderedProblems =
        RequestStructuralProblemOrder.order(List.copyOf(structuralProblems));
    Objects.requireNonNull(diagnosticRedactor, "diagnosticRedactor must not be null");
    Optional<RequestJsonNode> parsedRoot = Objects.requireNonNullElseGet(rawRoot, Optional::empty);
    List<RequestBindingFailure> discoveredFailures = new ArrayList<>(bindingFailures);
    Optional<WorkbookPlan> completePlan = Optional.empty();
    if (orderedProblems.isEmpty() && discoveredFailures.isEmpty()) {
      try {
        completePlan = boundFragments.completePlan();
      } catch (IllegalArgumentException exception) {
        discoveredFailures.add(
            RequestBindingFailureNormalizer.fromCompletePlan(exception, diagnosticRedactor));
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
    Optional<String> inferredPath =
        rawRoot.flatMap(
            root -> RequestBindingPathSupport.inferQualifiedInvariantFieldPath(root, failure));
    if (inferredPath.isPresent()) {
      String qualifiedPath = inferredPath.orElseThrow();
      return failure.rebasedAt(
          qualifiedPath,
          rawRoot.flatMap(
              root -> RequestJsonTokenLocationSupport.byteOffsetAt(root, qualifiedPath)));
    }
    Optional<Long> byteOffset =
        rawRoot.flatMap(
            root -> RequestJsonTokenLocationSupport.byteOffsetAt(root, failure.jsonPath()));
    return byteOffset.isPresent() ? failure.locatedAt(byteOffset) : failure;
  }
}
