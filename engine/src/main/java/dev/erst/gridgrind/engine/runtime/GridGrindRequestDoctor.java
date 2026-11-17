package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.plan.ExecutionModeSpec;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.GridGrindProblemDetail;
import dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces.RequestInput;
import dev.erst.gridgrind.core.problem.RequestDoctorReport;
import dev.erst.gridgrind.core.problem.RequestIntake;
import dev.erst.gridgrind.core.problem.RequestWarning;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Runs contract-derived linting for one authored request without mutating workbook sources. */
public final class GridGrindRequestDoctor {
  private final StaticRequestValidator staticValidator;
  private final NativeInvocationContext nativeContext;

  /** Creates the production doctor backed by the same request validator used for execution. */
  public GridGrindRequestDoctor(NativeInvocationContext context) {
    this(new StaticRequestValidator(), context);
  }

  GridGrindRequestDoctor(
      StaticRequestValidator staticValidator, NativeInvocationContext nativeContext) {
    this.nativeContext = Objects.requireNonNull(nativeContext);
    this.staticValidator =
        Objects.requireNonNull(staticValidator, "staticValidator must not be null");
  }

  /** Diagnoses tolerant authored input and closes the invocation before returning its report. */
  public RequestDoctorReport diagnose(RequestIntake analysis, RequestInput input) {
    return within(() -> diagnoseAnalysis(analysis, input));
  }

  /** Diagnoses authored input with host bindings and invocation-owned native preflight. */
  public RequestDoctorReport diagnose(
      RequestIntake analysis, RequestInput input, ExecutionInputBindings inputs) {
    return within(() -> diagnoseAnalysis(analysis, input, inputs));
  }

  /** Diagnoses a complete plan without source preflight. */
  public RequestDoctorReport diagnose(WorkbookPlan request) {
    return within(() -> diagnoseStatic(request));
  }

  /** Diagnoses a complete plan with host-bound source preflight. */
  public RequestDoctorReport diagnose(WorkbookPlan request, ExecutionInputBindings inputs) {
    return within(() -> diagnoseBound(request, inputs));
  }

  private RequestDoctorReport within(java.util.function.Supplier<RequestDoctorReport> work) {
    nativeContext.requireOpen();
    RequestDoctorReport report;
    try {
      report = work.get();
    } catch (RuntimeException | Error primary) {
      nativeContext.closePreserving(primary);
      throw primary;
    }
    try {
      nativeContext.close();
      return report;
    } catch (RuntimeException cleanup) {
      var problems = new java.util.ArrayList<>(report.problems());
      problems.add(
          GridGrindProblems.problem(
              dev.erst.gridgrind.core.problem.GridGrindProblemCode.INTERNAL_ERROR,
              "Native execution context cleanup failed.",
              new dev.erst.gridgrind.core.problem.ProblemContext.ExecuteRequest(
                  dev.erst.gridgrind.core.problem.ProblemContextRequestSurfaces.RequestShape
                      .unknown()),
              cleanup));
      return RequestDoctorReport.invalid(report.summary(), report.warnings(), problems);
    }
  }

  /** Returns all independently provable static findings from one tolerant request analysis. */
  private RequestDoctorReport diagnoseAnalysis(RequestIntake analysis, RequestInput requestInput) {
    Objects.requireNonNull(analysis, "analysis must not be null");
    Objects.requireNonNull(requestInput, "requestInput must not be null");
    List<GridGrindProblemDetail.Problem> problems =
        staticValidator.validate(analysis, requestInput);
    if (!problems.isEmpty()) {
      return invalidAnalysisReport(analysis, problems);
    }
    WorkbookPlan request = analysis.requireCompletePlan();
    return diagnoseStatic(request);
  }

  /** Returns static and source/input preflight findings from one tolerant request analysis. */
  private RequestDoctorReport diagnoseAnalysis(
      RequestIntake analysis, RequestInput requestInput, ExecutionInputBindings bindings) {
    Objects.requireNonNull(analysis, "analysis must not be null");
    Objects.requireNonNull(requestInput, "requestInput must not be null");
    Objects.requireNonNull(bindings, "bindings must not be null");
    List<GridGrindProblemDetail.Problem> problems =
        staticValidator.validate(analysis, requestInput);
    return analysis
        .completePlan()
        .map(request -> diagnose(request, Optional.of(bindings), Optional.of(analysis), problems))
        .orElseGet(() -> invalidAnalysisReport(analysis, problems));
  }

  /** Returns one machine-readable lint report for the supplied request. */
  private RequestDoctorReport diagnoseStatic(WorkbookPlan request) {
    return diagnose(Objects.requireNonNull(request, "request must not be null"), Optional.empty());
  }

  /**
   * Returns one machine-readable lint report for the supplied request using the provided authored
   * input bindings when input resolution should be validated as part of linting.
   */
  private RequestDoctorReport diagnoseBound(WorkbookPlan request, ExecutionInputBindings bindings) {
    return diagnose(
        Objects.requireNonNull(request, "request must not be null"),
        Optional.of(Objects.requireNonNull(bindings, "bindings must not be null")));
  }

  private RequestDoctorReport diagnose(
      WorkbookPlan request, Optional<ExecutionInputBindings> bindings) {
    return diagnose(request, bindings, Optional.empty(), staticValidator.validate(request));
  }

  private RequestDoctorReport diagnose(
      WorkbookPlan request,
      Optional<ExecutionInputBindings> bindings,
      Optional<RequestIntake> analysis,
      List<GridGrindProblemDetail.Problem> staticProblems) {
    RequestDoctorReport.Summary summary = summaryFor(request);
    List<RequestWarning> warnings =
        new java.util.ArrayList<>(GridGrindRequestWarnings.collect(request));
    List<GridGrindProblemDetail.Problem> problems = new java.util.ArrayList<>(staticProblems);
    if (bindings.isPresent() && staticProblems.isEmpty()) {
      ExecutionInputBindings boundInputs = bindings.orElseThrow().forRun();
      dev.erst.gridgrind.engine.backend.xssf.WorkbookTempScope.call(
          boundInputs.tempFileFactory()::createTempFile,
          () -> {
            try (ExecutionSession _ = boundInputs.session()) {
              RequestPreflight.Result preflight =
                  analysis
                      .map(value -> RequestPreflight.verify(request, boundInputs, value))
                      .orElseGet(() -> RequestPreflight.verify(request, boundInputs));
              problems.addAll(preflight.problems());
              warnings.addAll(preflight.warnings());
            } catch (java.io.IOException cleanup) {
              problems.add(
                  GridGrindProblems.fromException(
                      cleanup,
                      new dev.erst.gridgrind.core.problem.ProblemContext.ExecuteRequest(
                          ExecutionRequestPaths.requestShape(request))));
            }
            return Boolean.TRUE;
          });
    }
    if (!problems.isEmpty()) {
      return RequestDoctorReport.invalid(summary, warnings, problems);
    }
    if (!warnings.isEmpty()) {
      return RequestDoctorReport.warnings(summary, warnings);
    }
    return RequestDoctorReport.clean(summary);
  }

  private static RequestDoctorReport invalidAnalysisReport(
      RequestIntake analysis, List<GridGrindProblemDetail.Problem> problems) {
    return analysis
        .completePlan()
        .map(request ->
            RequestDoctorReport.invalid(
                summaryFor(request), GridGrindRequestWarnings.collect(request), problems))
        .orElseGet(() -> RequestDoctorReport.invalid(Optional.empty(), List.of(), problems));
  }

  private static RequestDoctorReport.Summary summaryFor(WorkbookPlan request) {
    Objects.requireNonNull(request, "request must not be null");
    ExecutionModeSpec executionMode = request.effectiveExecutionMode();
    WorkbookPlan.StepPartition stepPartition = request.stepPartition();
    int mutationStepCount = stepPartition.mutations().size();
    int assertionStepCount = stepPartition.assertions().size();
    int inspectionStepCount = stepPartition.inspections().size();
    return new RequestDoctorReport.Summary(
        ExecutionRequestPaths.reqSourceType(request),
        ExecutionRequestPaths.reqPersistenceType(request),
        executionMode.modeType(),
        request.calculationPolicy().effectiveStrategy().strategyType(),
        request.calculationPolicy().markRecalculateOnOpen(),
        SourceBackedPlanResolver.requiresStandardInput(request),
        request.steps().size(),
        mutationStepCount,
        assertionStepCount,
        inspectionStepCount);
  }
}
