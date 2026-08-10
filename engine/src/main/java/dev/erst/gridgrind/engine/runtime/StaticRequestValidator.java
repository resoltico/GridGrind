package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.DiagnosticOrder;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.json.RequestAnalysis;
import dev.erst.gridgrind.contract.json.RequestBoundFragments;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.contract.step.WorkbookOperationContracts;
import dev.erst.gridgrind.contract.step.WorkbookStep;
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
    fragments
        .steps()
        .ifPresent(
            steps ->
                steps.forEach(
                    step ->
                        step.value()
                            .flatMap(value -> targetViolation(value))
                            .ifPresent(
                                message ->
                                    problems.add(
                                        staticProblem(
                                            requestShape(fragments),
                                            java.util.Optional.of(analysis),
                                            "steps[" + step.index() + "].target.type",
                                            message)))));
    persistenceFailureMessage(fragments.source(), fragments.persistence())
        .ifPresent(
            message ->
                problems.add(
                    staticProblem(
                        requestShape(fragments),
                        java.util.Optional.of(analysis),
                        "persistence.type",
                        message)));
    if (fragments.execution().isPresent() && fragments.steps().isPresent()) {
      List<RequestBoundFragments.Step> steps = fragments.steps().orElseThrow();
      if (steps.stream().allMatch(step -> step.value().isPresent())) {
        addExecutionRuleProblems(
            problems,
            fragments.execution().orElseThrow(),
            fragments.source(),
            steps.stream().map(step -> step.value().orElseThrow()).toList(),
            requestShape(fragments),
            java.util.Optional.of(analysis));
      }
    }
    return DiagnosticOrder.problems(problems);
  }

  List<GridGrindProblemDetail.Problem> validate(WorkbookPlan request) {
    Objects.requireNonNull(request, "request must not be null");
    return DiagnosticOrder.problems(validatePlanRules(request, java.util.Optional.empty()));
  }

  private static List<GridGrindProblemDetail.Problem> validatePlanRules(
      WorkbookPlan request, java.util.Optional<RequestAnalysis> analysis) {
    List<GridGrindProblemDetail.Problem> problems = new ArrayList<>();
    for (int index = 0; index < request.steps().size(); index++) {
      int stepIndex = index;
      targetViolation(request.steps().get(index))
          .ifPresent(
              message ->
                  problems.add(
                      staticProblem(
                          ExecutionRequestPaths.requestShape(request),
                          analysis,
                          "steps[" + stepIndex + "].target.type",
                          message)));
    }
    addExecutionRuleProblems(
        problems,
        request.execution(),
        java.util.Optional.of(request.source()),
        request.steps(),
        ExecutionRequestPaths.requestShape(request),
        analysis);
    persistenceFailureMessage(
            java.util.Optional.of(request.source()), java.util.Optional.of(request.persistence()))
        .ifPresent(
            message ->
                problems.add(
                    staticProblem(
                        ExecutionRequestPaths.requestShape(request),
                        analysis,
                        "persistence.type",
                        message)));
    return List.copyOf(problems);
  }

  private static void addExecutionRuleProblems(
      List<GridGrindProblemDetail.Problem> problems,
      dev.erst.gridgrind.contract.dto.ExecutionPolicyInput execution,
      java.util.Optional<WorkbookPlan.WorkbookSource> source,
      List<WorkbookStep> steps,
      ProblemContextRequestSurfaces.RequestShape requestShape,
      java.util.Optional<RequestAnalysis> analysis) {
    for (String message :
        ExecutionModeRules.calculationPolicyFailures(execution.calculation(), steps)) {
      problems.add(staticProblem(requestShape, analysis, "execution.calculation", message));
    }
    for (String message :
        ExecutionModeRules.executionModeFailures(
            execution.mode(), execution.calculation(), source, steps)) {
      problems.add(staticProblem(requestShape, analysis, "execution.mode", message));
    }
  }

  private static java.util.Optional<String> targetViolation(WorkbookStep step) {
    return switch (step) {
      case MutationStep mutation ->
          WorkbookOperationContracts.targetViolation(mutation.action(), mutation.target());
      case AssertionStep assertion ->
          WorkbookOperationContracts.targetViolation(assertion.assertion(), assertion.target());
      case InspectionStep inspection ->
          WorkbookOperationContracts.targetViolation(inspection.query(), inspection.target());
    };
  }

  private static java.util.Optional<String> persistenceFailureMessage(
      java.util.Optional<WorkbookPlan.WorkbookSource> source,
      java.util.Optional<WorkbookPlan.WorkbookPersistence> persistence) {
    if (source.isEmpty() || persistence.isEmpty()) {
      return java.util.Optional.empty();
    }
    return switch (persistence.orElseThrow()) {
      case WorkbookPlan.WorkbookPersistence.Overwrite _ ->
          switch (source.orElseThrow()) {
            case WorkbookPlan.WorkbookSource.New _ ->
                java.util.Optional.of(
                    "OVERWRITE persistence requires an EXISTING source; a NEW workbook has no source file to overwrite");
            case WorkbookPlan.WorkbookSource.ExistingFile _ -> java.util.Optional.empty();
          };
      case WorkbookPlan.WorkbookPersistence.None _ -> java.util.Optional.empty();
      case WorkbookPlan.WorkbookPersistence.SaveAs _ -> java.util.Optional.empty();
    };
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
