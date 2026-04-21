package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import dev.erst.gridgrind.contract.dto.RequestWarning;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Runs contract-derived linting for one authored request without opening or mutating workbooks. */
public final class GridGrindRequestDoctor {
  private final ExecutionValidationSupport validationSupport;

  /** Creates the production doctor backed by the same request validator used for execution. */
  public GridGrindRequestDoctor() {
    this(new ExecutionValidationSupport());
  }

  GridGrindRequestDoctor(ExecutionValidationSupport validationSupport) {
    this.validationSupport =
        Objects.requireNonNull(validationSupport, "validationSupport must not be null");
  }

  /** Returns one machine-readable lint report for the supplied request. */
  public RequestDoctorReport diagnose(WorkbookPlan request) {
    if (request == null) {
      return RequestDoctorReport.invalid(
          null,
          List.of(),
          GridGrindProblems.problem(
              GridGrindProblemCode.INVALID_REQUEST,
              "request must not be null",
              new GridGrindResponse.ProblemContext.ValidateRequest(null, null),
              (Throwable) null));
    }

    RequestDoctorReport.Summary summary = summaryFor(request);
    List<RequestWarning> warnings = GridGrindRequestWarnings.collect(request);
    Optional<GridGrindResponse.Problem> validationProblem =
        validationSupport.validateRequest(request);
    if (validationProblem.isPresent()) {
      return RequestDoctorReport.invalid(summary, warnings, validationProblem.get());
    }
    if (!warnings.isEmpty()) {
      return RequestDoctorReport.warnings(summary, warnings);
    }
    return RequestDoctorReport.clean(summary);
  }

  private static RequestDoctorReport.Summary summaryFor(WorkbookPlan request) {
    Objects.requireNonNull(request, "request must not be null");
    ExecutionModeSelection executionModes = ExecutionModeRules.executionModes(request);
    int mutationStepCount =
        (int) request.steps().stream().filter(MutationStep.class::isInstance).count();
    int assertionStepCount =
        (int) request.steps().stream().filter(AssertionStep.class::isInstance).count();
    int inspectionStepCount =
        (int) request.steps().stream().filter(InspectionStep.class::isInstance).count();
    return new RequestDoctorReport.Summary(
        ExecutionRequestPaths.reqSourceType(request),
        ExecutionRequestPaths.reqPersistenceType(request),
        executionModes.readMode().name(),
        executionModes.writeMode().name(),
        request.calculationPolicy().effectiveStrategy().strategyType(),
        request.calculationPolicy().markRecalculateOnOpen(),
        SourceBackedPlanResolver.requiresStandardInput(request),
        request.steps().size(),
        mutationStepCount,
        assertionStepCount,
        inspectionStepCount);
  }
}
