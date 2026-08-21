package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.DiagnosticOrder;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.json.RequestAnalysis;
import dev.erst.gridgrind.contract.json.RequestBoundFragments;
import dev.erst.gridgrind.contract.step.WorkbookStaticRequest;
import dev.erst.gridgrind.contract.step.WorkbookStaticRequestContract;
import dev.erst.gridgrind.contract.step.WorkbookStaticStep;
import dev.erst.gridgrind.contract.step.WorkbookStaticViolation;
import dev.erst.gridgrind.engine.api.GridGrindProblems;
import dev.erst.gridgrind.engine.api.GridGrindRequestAnalysisProblems;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

/**
 * Single phase-three static validator for complete plans and tolerant partial request analyses.
 *
 * <p>A rule runs only after all of the fragments it reads are bound. This intentionally retains
 * valid sibling steps while suppressing speculative secondary diagnostics beneath malformed
 * fragments.
 */
final class StaticRequestValidator {
  List<GridGrindProblemDetail.Problem> validate(
      RequestAnalysis analysis, ProblemContextRequestSurfaces.RequestInput requestInput) {
    Objects.requireNonNull(analysis, "analysis must not be null");
    Objects.requireNonNull(requestInput, "requestInput must not be null");
    List<GridGrindProblemDetail.Problem> problems =
        new ArrayList<>(GridGrindRequestAnalysisProblems.project(analysis, requestInput));
    RequestBoundFragments fragments = analysis.boundFragments();
    addContractViolations(
        problems,
        partialRequest(fragments),
        requestShape(fragments),
        java.util.Optional.of(analysis));
    return DiagnosticOrder.problems(problems);
  }

  List<GridGrindProblemDetail.Problem> validate(WorkbookPlan request) {
    Objects.requireNonNull(request, "request must not be null");
    return DiagnosticOrder.problems(validatePlanRules(request, java.util.Optional.empty()));
  }

  private static List<GridGrindProblemDetail.Problem> validatePlanRules(
      WorkbookPlan request, java.util.Optional<RequestAnalysis> analysis) {
    List<GridGrindProblemDetail.Problem> problems = new ArrayList<>();
    addContractViolations(
        problems,
        WorkbookStaticRequestContract.from(request),
        ExecutionRequestPaths.requestShape(request),
        analysis);
    return List.copyOf(problems);
  }

  private static void addContractViolations(
      List<GridGrindProblemDetail.Problem> problems,
      WorkbookStaticRequest request,
      ProblemContextRequestSurfaces.RequestShape requestShape,
      java.util.Optional<RequestAnalysis> analysis) {
    for (WorkbookStaticViolation violation : WorkbookStaticRequestContract.validate(request)) {
      problems.add(
          staticProblem(requestShape, analysis, violation.jsonPath(), violation.message()));
    }
  }

  private static WorkbookStaticRequest partialRequest(RequestBoundFragments fragments) {
    return new WorkbookStaticRequest(
        fragments.source(),
        fragments.persistence(),
        fragments.execution(),
        fragments.steps().orElseGet(List::of).stream()
            .map(step -> new WorkbookStaticStep(step.index(), step.value()))
            .toList());
  }

  private static GridGrindProblemDetail.Problem staticProblem(
      ProblemContextRequestSurfaces.RequestShape requestShape,
      java.util.Optional<RequestAnalysis> analysis,
      String jsonPath,
      String message) {
    ProblemContext.ValidateRequest context = new ProblemContext.ValidateRequest(requestShape);
    if (analysis.isPresent()) {
      context = context.withJson(analysis.orElseThrow().jsonLocationAt(jsonPath));
    }
    return GridGrindProblems.problem(
        GridGrindProblemCode.INVALID_REQUEST, message, context, List.of());
  }

  private static ProblemContextRequestSurfaces.RequestShape requestShape(
      RequestBoundFragments fragments) {
    return requestShape(fragments.source(), fragments.persistence());
  }

  private static ProblemContextRequestSurfaces.RequestShape requestShape(
      java.util.Optional<WorkbookPlan.WorkbookSource> source,
      java.util.Optional<WorkbookPlan.WorkbookPersistence> persistence) {
    if (source.isEmpty() || persistence.isEmpty()) {
      return ProblemContextRequestSurfaces.RequestShape.unknown();
    }
    String sourceType =
        source.orElseThrow() instanceof WorkbookPlan.WorkbookSource.New ? "NEW" : "EXISTING";
    String persistenceType =
        switch (persistence.orElseThrow()) {
          case WorkbookPlan.WorkbookPersistence.None _ -> "NONE";
          case WorkbookPlan.WorkbookPersistence.SaveAs _ -> "SAVE_AS";
          case WorkbookPlan.WorkbookPersistence.Overwrite _ -> "OVERWRITE";
        };
    return ProblemContextRequestSurfaces.RequestShape.known(sourceType, persistenceType);
  }
}
