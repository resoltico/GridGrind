package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.plan.ResolvedPlan;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.plan.WorkbookStep;
import dev.erst.gridgrind.core.problem.DiagnosticOrder;
import dev.erst.gridgrind.core.problem.GridGrindProblemDetail;
import dev.erst.gridgrind.core.problem.ProblemContext;
import dev.erst.gridgrind.core.problem.ProblemContextWorkbookSurfaces.InputReference;
import dev.erst.gridgrind.core.problem.RequestIntake;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/**
 * Sole factory for engine-only prepared execution capabilities, shared by doctor and executor.
 *
 * <p>Static admission precedes all resource acquisition. After static and authority checks, input
 * resolution and source verification collect independent findings. Prepared native resources remain
 * session-owned until an executing workflow takes their exclusive ownership.
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
      WorkbookPlan request, ExecutionInputBindings bindings, RequestIntake analysis) {
    return verify(
        request,
        bindings,
        Optional.of(Objects.requireNonNull(analysis, "analysis must not be null")),
        SourceBackedPlanResolver::resolveStep);
  }

  @SuppressWarnings("PMD.CloseResource")
  // Path and prepared-source owners are registered with the session, including refusal paths.
  private static Result verify(
      WorkbookPlan request,
      ExecutionInputBindings bindings,
      Optional<RequestIntake> analysis,
      StepResolver stepResolver) {
    Objects.requireNonNull(request, "request must not be null");
    Objects.requireNonNull(bindings, "bindings must not be null");
    Objects.requireNonNull(analysis, "analysis must not be null");
    Objects.requireNonNull(stepResolver, "stepResolver must not be null");
    List<GridGrindProblemDetail.Problem> staticProblems =
        new StaticRequestValidator().validate(request);
    if (!staticProblems.isEmpty()) {
      return new Result(Optional.empty(), staticProblems, List.of());
    }
    List<GridGrindProblemDetail.Problem> problems = new ArrayList<>();
    RequestPathAccess pathAccess =
        new RequestPathAccess(
            bindings.workingDirectory(),
            bindings.tempFileFactory(),
            bindings.executionGrant(),
            bindings.snapshots(),
            bindings.publicationFacts());
    bindings.own(pathAccess::close);
    ExecutionInputBindings locatedBindings =
        bindings
            .withRequestPathAccess(pathAccess)
            .withInputResolutionOrigins(InputResolutionOrigins.forRequest(request, analysis));
    RequestPreflightStaticPaths.validate(request, locatedBindings, problems);
    if (problems.isEmpty()) {
      addGrantViolations(request, locatedBindings, problems);
    }
    if (!problems.isEmpty()) {
      return new Result(
          Optional.empty(), DiagnosticOrder.problems(problems), pathAccess.warnings());
    }
    Optional<ResolvedPlan> resolvedRequest =
        RequestPreflightInputResolver.resolve(request, locatedBindings, stepResolver, problems);
    HostAcceptanceInputResolver.resolve(request, locatedBindings, problems);
    RequestPreflightPaths.verify(
        request, locatedBindings.withInputResolutionScope("persistence"), problems);
    PreparedWorkbookSource source =
        RequestPreflightWorkbookSource.prepare(request, locatedBindings, problems);
    if (resolvedRequest.isPresent() && source.dateSystem().isPresent())
      TemporalInputValidation.validate(
          resolvedRequest.orElseThrow(), source.dateSystem().orElseThrow(), analysis, problems);
    if (!problems.isEmpty()) {
      return new Result(
          Optional.empty(), DiagnosticOrder.problems(problems), pathAccess.warnings());
    }
    locatedBindings.recordExecutionStrategy(source.strategy());
    PreparedRequest prepared =
        new PreparedRequest(resolvedRequest.orElseThrow(), locatedBindings, source);
    bindings.session().prepared(prepared);
    return new Result(Optional.of(prepared), List.of(), pathAccess.warnings());
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
        .map(location ->
            new ProblemContext.ResolveInputs(
                context.request(), context.input(), Optional.of(location)))
        .orElse(context);
  }

  private static void addGrantViolations(
      WorkbookPlan request,
      ExecutionInputBindings bindings,
      List<GridGrindProblemDetail.Problem> problems) {
    try {
      ExecutionGrantValidator.violations(
              request, bindings.executionGrant(), bindings.requestPathAccess().pathIdentity())
          .forEach(violation ->
              problems.add(
                  GridGrindProblems.fromException(
                      violation,
                      new ProblemContext.ValidateRequest(
                          ExecutionRequestPaths.requestShape(request)))));
    } catch (IOException | RuntimeException failure) {
      problems.add(
          GridGrindProblems.fromException(
              failure,
              new ProblemContext.ValidateRequest(ExecutionRequestPaths.requestShape(request))));
    }
    HostAcceptancePolicyValidator.violations(request, bindings.executionGrant())
        .forEach(violation ->
            problems.add(
                GridGrindProblems.fromException(
                    violation,
                    new ProblemContext.ValidateRequest(
                        ExecutionRequestPaths.requestShape(request)))));
  }

  static ProblemContext.OpenWorkbook openWorkbookContext(
      WorkbookPlan request, ExecutionInputBindings bindings) {
    return new ProblemContext.OpenWorkbook(
        ExecutionRequestPaths.requestShape(request),
        ExecutionRequestPaths.workbookReference(request, bindings.workingDirectory()),
        java.util.Optional.empty());
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
      List<dev.erst.gridgrind.core.problem.RequestWarning> warnings) {
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

    ResolvedPlan preparedPlan() {
      return requirePreparedRequest().plan();
    }

    ExecutionInputBindings preparedBindings() {
      return requirePreparedRequest().bindings();
    }
  }

  /**
   * Engine-only capability issued after static, authority, operand and source preparation succeed.
   */
  static final class PreparedRequest implements AutoCloseable {
    private final PreparedWorkbookSource source;
    private final ResolvedPlan plan;
    private final ExecutionInputBindings bindings;

    private PreparedRequest(
        ResolvedPlan plan, ExecutionInputBindings bindings, PreparedWorkbookSource source) {
      this.source = Objects.requireNonNull(source);
      this.plan = Objects.requireNonNull(plan);
      this.bindings = Objects.requireNonNull(bindings);
    }

    PreparedWorkbookSource source() {
      return source;
    }

    ResolvedPlan plan() {
      return plan;
    }

    ExecutionInputBindings bindings() {
      return bindings;
    }

    @Override
    public void close() throws IOException {
      bindings.closeRunResources();
    }
  }
}
