package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.plan.PreparedStep;
import dev.erst.gridgrind.core.plan.ResolvedPlan;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.plan.WorkbookStep;
import dev.erst.gridgrind.core.problem.GridGrindProblemDetail;
import dev.erst.gridgrind.engine.api.GridGrindProblems;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

/** Resolves every source-backed step input while retaining independent sibling failures. */
final class RequestPreflightInputResolver {
  private RequestPreflightInputResolver() {}

  static Optional<ResolvedPlan> resolve(
      WorkbookPlan request,
      ExecutionInputBindings bindings,
      RequestPreflight.StepResolver stepResolver,
      List<GridGrindProblemDetail.Problem> problems) {
    List<PreparedStep> resolvedSteps = new ArrayList<>(request.steps().size());
    boolean failed = false;
    for (int index = 0; index < request.steps().size(); index++) {
      WorkbookStep step = request.steps().get(index);
      ExecutionInputBindings occurrence =
          bindings.withInputResolutionScope("steps[" + index + "]").withNewMediaUse();
      try {
        WorkbookStep bound = stepResolver.resolve(step, occurrence);
        occurrence.mediaUse().reserveBound(step, bound, occurrence, index);
        PreparedStep prepared = PreparedStep.prepare(step, bound);
        occurrence.mediaUse().inspect(prepared);
        resolvedSteps.add(prepared);
      } catch (InputResolutionBatchException exception) {
        addBatchFailures(request, exception, problems);
        failed = true;
      } catch (Exception exception) {
        problems.add(
            GridGrindProblems.fromException(
                exception,
                RequestPreflight.resolveInputsContext(
                    request, occurrence.mediaUse().failure(exception))));
        failed = true;
      }
    }
    if (failed) {
      return Optional.empty();
    }
    return Optional.of(new ResolvedPlan(request, resolvedSteps));
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
