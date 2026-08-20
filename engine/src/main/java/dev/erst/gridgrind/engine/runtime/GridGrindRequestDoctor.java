package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.ProblemContextRequestSurfaces.RequestInput;
import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import dev.erst.gridgrind.contract.dto.RequestWarning;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.json.RequestAnalysis;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Runs contract-derived linting for one authored request without mutating workbook sources. */
public final class GridGrindRequestDoctor {
  private final StaticRequestValidator staticValidator;

  /** Creates the production doctor backed by the same request validator used for execution. */
  public GridGrindRequestDoctor() {
    this(new StaticRequestValidator());
  }

  GridGrindRequestDoctor(StaticRequestValidator staticValidator) {
    this.staticValidator =
        Objects.requireNonNull(staticValidator, "staticValidator must not be null");
  }

  /** Returns all independently provable static findings from one tolerant request analysis. */
  public RequestDoctorReport diagnose(RequestAnalysis analysis, RequestInput requestInput) {
    Objects.requireNonNull(analysis, "analysis must not be null");
    Objects.requireNonNull(requestInput, "requestInput must not be null");
    List<GridGrindProblemDetail.Problem> problems =
        staticValidator.validate(analysis, requestInput);
    if (!problems.isEmpty()) {
      return invalidAnalysisReport(analysis, problems);
    }
    WorkbookPlan request = analysis.requireCompletePlan();
    return diagnose(request);
  }

  /** Returns static and source/input preflight findings from one tolerant request analysis. */
  public RequestDoctorReport diagnose(
      RequestAnalysis analysis, RequestInput requestInput, ExecutionInputBindings bindings) {
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
  public RequestDoctorReport diagnose(WorkbookPlan request) {
    return diagnose(Objects.requireNonNull(request, "request must not be null"), Optional.empty());
  }

  /**
   * Returns one machine-readable lint report for the supplied request using the provided authored
   * input bindings when input resolution should be validated as part of linting.
   */
  public RequestDoctorReport diagnose(WorkbookPlan request, ExecutionInputBindings bindings) {
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
      Optional<RequestAnalysis> analysis,
      List<GridGrindProblemDetail.Problem> staticProblems) {
    RequestDoctorReport.Summary summary = summaryFor(request);
    List<RequestWarning> warnings = GridGrindRequestWarnings.collect(request);
    List<GridGrindProblemDetail.Problem> problems = new java.util.ArrayList<>(staticProblems);
    if (bindings.isPresent()) {
      ExecutionInputBindings boundInputs = bindings.orElseThrow();
      RequestPreflight.Result preflight =
          analysis
              .map(value -> RequestPreflight.verify(request, boundInputs, value))
              .orElseGet(() -> RequestPreflight.verify(request, boundInputs));
      problems.addAll(preflight.problems());
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
      RequestAnalysis analysis, List<GridGrindProblemDetail.Problem> problems) {
    return analysis
        .completePlan()
        .map(
            request ->
                RequestDoctorReport.invalid(
                    summaryFor(request), GridGrindRequestWarnings.collect(request), problems))
        .orElseGet(() -> RequestDoctorReport.invalid(Optional.empty(), List.of(), problems));
  }

  private static RequestDoctorReport.Summary summaryFor(WorkbookPlan request) {
    Objects.requireNonNull(request, "request must not be null");
    ExecutionModeInput executionMode = ExecutionWorkflowRouting.executionMode(request);
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
