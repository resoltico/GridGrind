package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.step.WorkbookStep;
import dev.erst.gridgrind.engine.api.GridGrindProblems;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

/** Resolves every source-backed step input while retaining independent sibling failures. */
final class RequestPreflightInputResolver {
  private RequestPreflightInputResolver() {}

  static Optional<WorkbookPlan> resolve(
      WorkbookPlan request,
      ExecutionInputBindings bindings,
      RequestPreflight.StepResolver stepResolver,
      List<GridGrindProblemDetail.Problem> problems) {
    List<WorkbookStep> resolvedSteps = new ArrayList<>(request.steps().size());
    boolean failed = false;
    for (WorkbookStep step : request.steps()) {
      try {
        resolvedSteps.add(stepResolver.resolve(step, bindings));
      } catch (InputResolutionBatchException exception) {
        addBatchFailures(request, exception, problems);
        resolvedSteps.add(step);
        failed = true;
      } catch (Exception exception) {
        problems.add(
            GridGrindProblems.fromException(
                exception,
                RequestPreflight.resolveInputsContext(
                    request, InputResolutionFailure.unlocated(exception))));
        resolvedSteps.add(step);
        failed = true;
      }
    }
    if (failed) {
      return Optional.empty();
    }
    return Optional.of(
        new WorkbookPlan(
            request.protocolVersion(),
            request.planId(),
            request.source(),
            request.persistence(),
            request.execution(),
            request.formulaEnvironment(),
            resolvedSteps));
  }

  private static void addBatchFailures(
      WorkbookPlan request,
      InputResolutionBatchException exception,
      List<GridGrindProblemDetail.Problem> problems) {
    for (InputResolutionFailure failure : exception.failures()) {
      problems.add(
          GridGrindProblems.fromException(
              failure.exception(), RequestPreflight.resolveInputsContext(request, failure)));
    }
  }
}
