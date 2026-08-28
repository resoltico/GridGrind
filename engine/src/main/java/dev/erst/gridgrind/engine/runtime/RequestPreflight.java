package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.DiagnosticOrder;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.ProblemContext;
import dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.InputReference;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.json.RequestAnalysis;
import dev.erst.gridgrind.contract.step.WorkbookStep;
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

  @SuppressWarnings({"PMD.CloseResource", "PMD.UseTryWithResources"})
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
    RequestPathAccess pathAccess =
        new RequestPathAccess(bindings.workingDirectory(), bindings.tempFileFactory());
    boolean prepared = false;
    try {
      ExecutionInputBindings locatedBindings =
          bindings
              .withRequestPathAccess(pathAccess)
              .withInputResolutionOrigins(InputResolutionOrigins.forRequest(request, analysis));
      Optional<WorkbookPlan> resolvedRequest =
          RequestPreflightInputResolver.resolve(request, locatedBindings, stepResolver, problems);
      RequestPreflightPaths.verify(request, locatedBindings, problems);
      RequestPreflightWorkbookSource.verify(request, locatedBindings, problems);
      Result result =
          new Result(
              problems.isEmpty()
                  ? resolvedRequest.map(value -> new PreparedRequest(value, locatedBindings))
                  : Optional.empty(),
              DiagnosticOrder.problems(problems),
              pathAccess.warnings());
      prepared = result.preparedRequest().isPresent();
      return result;
    } finally {
      if (!prepared) {
        closeQuietly(pathAccess);
      }
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
                : failure.input());
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
      Optional<PreparedRequest> preparedRequest,
      List<GridGrindProblemDetail.Problem> problems,
      List<dev.erst.gridgrind.contract.dto.RequestWarning> warnings) {
    Result {
      Objects.requireNonNull(preparedRequest, "preparedRequest must not be null");
      Objects.requireNonNull(problems, "problems must not be null");
      Objects.requireNonNull(warnings, "warnings must not be null");
      problems = List.copyOf(problems);
      warnings = List.copyOf(warnings);
      if (preparedRequest.isPresent() != problems.isEmpty()) {
        throw new IllegalArgumentException(
            "preparedRequest must be present exactly when phase-four problems are empty");
      }
    }

    PreparedRequest requirePreparedRequest() {
      return preparedRequest.orElseThrow(
          () -> new IllegalStateException("A failed preflight cannot produce a resolved request"));
    }

    WorkbookPlan preparedPlan() {
      return requirePreparedRequest().request();
    }

    ExecutionInputBindings preparedBindings() {
      return requirePreparedRequest().bindings();
    }

    /** Releases request-private resources retained by one successful preflight result. */
    void release() {
      preparedRequest.ifPresent(RequestPreflight::closeQuietly);
    }
  }

  record PreparedRequest(WorkbookPlan request, ExecutionInputBindings bindings)
      implements AutoCloseable {
    PreparedRequest {
      Objects.requireNonNull(request, "request must not be null");
      Objects.requireNonNull(bindings, "bindings must not be null");
    }

    @Override
    public void close() throws IOException {
      bindings.requestPathAccess().close();
    }
  }

  static void closeQuietly(AutoCloseable closeable) {
    try {
      closeable.close();
    } catch (Exception ignored) {
      // The original preflight finding is always more actionable than private-resource cleanup.
    }
  }
}
