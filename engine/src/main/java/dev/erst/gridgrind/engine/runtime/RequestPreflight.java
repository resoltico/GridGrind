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
          resolveInputs(request, locatedBindings, stepResolver, problems);
      preflightRequestOwnedPaths(request, locatedBindings, problems);
      preflightWorkbookSource(request, locatedBindings, problems);
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
    try (ExcelWorkbook workbook = workbookSupport.openWorkbook(request.source(), null, bindings)) {
      Objects.requireNonNull(workbook, "workbook must not be null");
    } catch (Exception exception) {
      problems.add(GridGrindProblems.fromException(exception, context));
    }
  }

  private static void preflightRequestOwnedPaths(
      WorkbookPlan request,
      ExecutionInputBindings bindings,
      List<GridGrindProblemDetail.Problem> problems) {
    preflightFormulaEnvironmentPaths(request, bindings, problems);
    preflightPersistenceMaterial(request, bindings, problems);
    preflightPersistenceTarget(request, bindings, problems);
  }

  private static void preflightFormulaEnvironmentPaths(
      WorkbookPlan request,
      ExecutionInputBindings bindings,
      List<GridGrindProblemDetail.Problem> problems) {
    for (var externalWorkbook : request.formulaEnvironment().externalWorkbooks()) {
      try {
        bindings
            .requestPathAccess()
            .materializeRead(
                externalWorkbook.path(),
                "formulaEnvironment.externalWorkbooks",
                "gridgrind-formula-workbook-",
                ".xlsx");
      } catch (IOException exception) {
        addFormulaPathProblem(
            request,
            externalWorkbook,
            SourceBackedPlanResolver.inputFileFailure(
                externalWorkbook.path(), "formula external workbook", exception),
            problems);
      } catch (RuntimeException exception) {
        addFormulaPathProblem(request, externalWorkbook, exception, problems);
      }
    }
  }

  private static void addFormulaPathProblem(
      WorkbookPlan request,
      dev.erst.gridgrind.contract.dto.FormulaExternalWorkbookInput externalWorkbook,
      Exception exception,
      List<GridGrindProblemDetail.Problem> problems) {
    problems.add(
        GridGrindProblems.fromException(
            exception,
            new ProblemContext.ResolveInputs(
                ExecutionRequestPaths.requestShape(request),
                InputReference.path("formula external workbook", externalWorkbook.path()))));
  }

  private static void preflightPersistenceMaterial(
      WorkbookPlan request,
      ExecutionInputBindings bindings,
      List<GridGrindProblemDetail.Problem> problems) {
    try {
      ExecutionRequestPaths.persistenceOptions(request.persistence(), bindings);
    } catch (Exception exception) {
      problems.add(
          GridGrindProblems.fromException(
              exception,
              new ProblemContext.PersistWorkbook(
                  ExecutionRequestPaths.requestShape(request),
                  ExecutionRequestPaths.persistenceReference(
                      request, bindings.workingDirectory()))));
    }
  }

  private static void preflightPersistenceTarget(
      WorkbookPlan request,
      ExecutionInputBindings bindings,
      List<GridGrindProblemDetail.Problem> problems) {
    String persistencePath;
    dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition disposition;
    switch (request.persistence()) {
      case WorkbookPlan.WorkbookPersistence.None _ -> {
        return;
      }
      case WorkbookPlan.WorkbookPersistence.SaveAs saveAs -> {
        persistencePath = saveAs.path();
        disposition = ExecutionRequestPaths.writeDisposition(saveAs);
      }
      case WorkbookPlan.WorkbookPersistence.Overwrite _ -> {
        if (!(request.source() instanceof WorkbookPlan.WorkbookSource.ExistingFile existingFile)) {
          return;
        }
        persistencePath = existingFile.path();
        disposition = dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING;
      }
    }
    try {
      bindings.requestPathAccess().prepareOutput(persistencePath, "persistence", disposition);
    } catch (Exception exception) {
      problems.add(
          GridGrindProblems.fromException(
              exception,
              new ProblemContext.PersistWorkbook(
                  ExecutionRequestPaths.requestShape(request),
                  ExecutionRequestPaths.persistenceReference(
                      request, bindings.workingDirectory()))));
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
