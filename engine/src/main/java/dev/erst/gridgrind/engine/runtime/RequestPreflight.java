package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.DiagnosticOrder;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.InputReference;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.json.RequestAnalysis;
import dev.erst.gridgrind.contract.step.WorkbookStep;
import dev.erst.gridgrind.engine.api.GridGrindProblems;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/**
 * Shared non-mutating phase-four request preflight for doctor and executor paths.
 *
 * <p>Input resolution and opening an existing source are independent checks: an unreadable asset
 * must not hide a bad source password or an inaccessible workbook. The preflight opens only to
 * prove availability, closes immediately, and never reaches workbook mutation or persistence.
 */
final class RequestPreflight {
  private RequestPreflight() {}

  static Result verify(WorkbookPlan request, ExecutionInputBindings bindings) {
    return verify(request, bindings, Optional.empty(), SourceBackedPlanResolver::resolveStep);
  }

  static Result verify(
      WorkbookPlan request, ExecutionInputBindings bindings, StepResolver stepResolver) {
    return verify(request, bindings, Optional.empty(), stepResolver);
  }

  static Result verify(
      WorkbookPlan request, ExecutionInputBindings bindings, RequestAnalysis analysis) {
    return verify(
        request,
        bindings,
        Optional.of(Objects.requireNonNull(analysis, "analysis must not be null")),
        SourceBackedPlanResolver::resolveStep);
  }

  private static Result verify(
      WorkbookPlan request,
      ExecutionInputBindings bindings,
      Optional<RequestAnalysis> analysis,
      StepResolver stepResolver) {
    Objects.requireNonNull(request, "request must not be null");
    Objects.requireNonNull(bindings, "bindings must not be null");
    Objects.requireNonNull(analysis, "analysis must not be null");
    Objects.requireNonNull(stepResolver, "stepResolver must not be null");
    List<GridGrindProblemDetail.Problem> problems = new ArrayList<>();
    ExecutionInputBindings locatedBindings =
        bindings.withInputResolutionOrigins(InputResolutionOrigins.forRequest(request, analysis));
    Optional<WorkbookPlan> resolvedRequest =
        resolveInputs(request, locatedBindings, stepResolver, problems);
    preflightWorkbookSource(request, locatedBindings, problems);
    return new Result(
        problems.isEmpty() ? resolvedRequest : Optional.empty(),
        DiagnosticOrder.problems(problems));
  }

  private static Optional<WorkbookPlan> resolveInputs(
      WorkbookPlan request,
      ExecutionInputBindings bindings,
      StepResolver stepResolver,
      List<GridGrindProblemDetail.Problem> problems) {
    List<WorkbookStep> resolvedSteps = new ArrayList<>(request.steps().size());
    boolean failed = false;
    for (WorkbookStep step : request.steps()) {
      try {
        resolvedSteps.add(stepResolver.resolve(step, bindings));
      } catch (InputResolutionBatchException exception) {
        for (InputResolutionFailure failure : exception.failures()) {
          problems.add(
              GridGrindProblems.fromException(
                  failure.exception(), resolveInputsContext(request, failure)));
        }
        resolvedSteps.add(step);
        failed = true;
      } catch (Exception exception) {
        // One malformed external asset must not hide unrelated source-backed sibling inputs.
        problems.add(
            GridGrindProblems.fromException(
                exception,
                resolveInputsContext(request, InputResolutionFailure.unlocated(exception))));
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

  private static void preflightWorkbookSource(
      WorkbookPlan request,
      ExecutionInputBindings bindings,
      List<GridGrindProblemDetail.Problem> problems) {
    if (!(request.source() instanceof WorkbookPlan.WorkbookSource.ExistingFile)) {
      return;
    }
    ProblemContext.OpenWorkbook context = openWorkbookContext(request, bindings);
    ExecutionWorkbookSupport workbookSupport =
        new ExecutionWorkbookSupport(bindings.tempFileFactory());
    try (ExcelWorkbook workbook =
        workbookSupport.openWorkbook(
            request.source(), request.formulaEnvironment(), bindings.workingDirectory())) {
      Objects.requireNonNull(workbook, "workbook must not be null");
    } catch (Exception exception) {
      problems.add(GridGrindProblems.fromException(exception, context));
    }
  }

  static ProblemContext.ResolveInputs resolveInputsContext(
      WorkbookPlan request, InputResolutionFailure failure) {
    Exception exception = failure.exception();
    ProblemContext.ResolveInputs context =
        new ProblemContext.ResolveInputs(
            ExecutionRequestPaths.requestShape(request),
            exception instanceof InputSourceException inputSourceException
                ? inputSourceException.inputPath() != null
                    ? InputReference.path(
                        inputSourceException.inputKind(), inputSourceException.inputPath())
                    : InputReference.kind(inputSourceException.inputKind())
                : InputReference.unknown());
    return failure
        .json()
        .map(
            location ->
                new ProblemContext.ResolveInputs(
                    context.request(), context.input(), Optional.of(location)))
        .orElse(context);
  }

  static ProblemContext.OpenWorkbook openWorkbookContext(
      WorkbookPlan request, ExecutionInputBindings bindings) {
    return new ProblemContext.OpenWorkbook(
        ExecutionRequestPaths.requestShape(request),
        ExecutionRequestPaths.workbookReference(request, bindings.workingDirectory()));
  }

  /** Resolves one step's externally authored input without mutating a workbook. */
  @FunctionalInterface
  interface StepResolver {
    /** Resolves the supplied step from the invocation bindings. */
    WorkbookStep resolve(WorkbookStep step, ExecutionInputBindings bindings) throws IOException;
  }

  /** One phase-four result carrying a resolved request only when every prerequisite succeeded. */
  record Result(
      Optional<WorkbookPlan> resolvedRequest, List<GridGrindProblemDetail.Problem> problems) {
    Result {
      Objects.requireNonNull(resolvedRequest, "resolvedRequest must not be null");
      Objects.requireNonNull(problems, "problems must not be null");
      problems = List.copyOf(problems);
      if (resolvedRequest.isPresent() != problems.isEmpty()) {
        throw new IllegalArgumentException(
            "resolvedRequest must be present exactly when phase-four problems are empty");
      }
    }

    WorkbookPlan requireResolvedRequest() {
      return resolvedRequest.orElseThrow(
          () -> new IllegalStateException("A failed preflight cannot produce a resolved request"));
    }
  }
}
