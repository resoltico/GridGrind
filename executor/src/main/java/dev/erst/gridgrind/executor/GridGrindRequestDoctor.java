package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.RequestDoctorReport;
import dev.erst.gridgrind.contract.dto.RequestWarning;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Runs contract-derived linting for one authored request without mutating workbook sources. */
public final class GridGrindRequestDoctor {
  private final ExecutionValidationSupport validationSupport;
  private final ExecutionWorkbookSupport workbookSupport;

  /** Creates the production doctor backed by the same request validator used for execution. */
  public GridGrindRequestDoctor() {
    this(new ExecutionValidationSupport(), defaultWorkbookSupport());
  }

  GridGrindRequestDoctor(ExecutionValidationSupport validationSupport) {
    this(validationSupport, defaultWorkbookSupport());
  }

  GridGrindRequestDoctor(
      ExecutionValidationSupport validationSupport, ExecutionWorkbookSupport workbookSupport) {
    this.validationSupport =
        Objects.requireNonNull(validationSupport, "validationSupport must not be null");
    this.workbookSupport =
        Objects.requireNonNull(workbookSupport, "workbookSupport must not be null");
  }

  /** Returns one machine-readable lint report for the supplied request. */
  public RequestDoctorReport diagnose(WorkbookPlan request) {
    return diagnose(request, Optional.empty());
  }

  /**
   * Returns one machine-readable lint report for the supplied request using the provided authored
   * input bindings when input resolution should be validated as part of linting.
   */
  public RequestDoctorReport diagnose(WorkbookPlan request, ExecutionInputBindings bindings) {
    return diagnose(
        request, Optional.of(Objects.requireNonNull(bindings, "bindings must not be null")));
  }

  private RequestDoctorReport diagnose(
      WorkbookPlan request, Optional<ExecutionInputBindings> bindings) {
    if (request == null) {
      return RequestDoctorReport.invalid(
          Optional.empty(),
          List.of(),
          GridGrindProblems.problem(
              GridGrindProblemCode.INVALID_REQUEST,
              "request must not be null",
              new dev.erst.gridgrind.contract.dto.ProblemContext.ValidateRequest(
                  dev.erst.gridgrind.contract.dto.ProblemContext.RequestShape.unknown()),
              (Throwable) null));
    }

    RequestDoctorReport.Summary summary = summaryFor(request);
    List<RequestWarning> warnings = GridGrindRequestWarnings.collect(request);
    Optional<GridGrindResponse.Problem> validationProblem =
        validationSupport.validateRequest(request);
    if (validationProblem.isPresent()) {
      return RequestDoctorReport.invalid(summary, warnings, validationProblem.get());
    }
    if (bindings.isPresent()) {
      ExecutionInputBindings boundInputs = bindings.orElseThrow();
      WorkbookPlan resolvedRequest;
      try {
        resolvedRequest = SourceBackedPlanResolver.resolve(request, boundInputs);
      } catch (Exception exception) {
        return RequestDoctorReport.invalid(
            summary,
            warnings,
            GridGrindProblems.fromException(exception, resolveInputsContext(request, exception)));
      }
      Optional<GridGrindResponse.Problem> openProblem =
          preflightWorkbookSource(resolvedRequest, boundInputs);
      if (openProblem.isPresent()) {
        return RequestDoctorReport.invalid(summary, warnings, openProblem.get());
      }
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

  private static dev.erst.gridgrind.contract.dto.ProblemContext.ResolveInputs resolveInputsContext(
      WorkbookPlan request, Exception exception) {
    return new dev.erst.gridgrind.contract.dto.ProblemContext.ResolveInputs(
        ExecutionRequestPaths.requestShape(request),
        exception instanceof InputSourceException inputSourceException
            ? inputSourceException.inputPath() != null
                ? dev.erst.gridgrind.contract.dto.ProblemContext.InputReference.path(
                    inputSourceException.inputKind(), inputSourceException.inputPath())
                : dev.erst.gridgrind.contract.dto.ProblemContext.InputReference.kind(
                    inputSourceException.inputKind())
            : dev.erst.gridgrind.contract.dto.ProblemContext.InputReference.unknown());
  }

  private Optional<GridGrindResponse.Problem> preflightWorkbookSource(
      WorkbookPlan request, ExecutionInputBindings bindings) {
    if (!(request.source() instanceof WorkbookPlan.WorkbookSource.ExistingFile)) {
      return Optional.empty();
    }
    dev.erst.gridgrind.contract.dto.ProblemContext.OpenWorkbook context =
        openWorkbookContext(request, bindings);
    try (ExcelWorkbook workbook =
        workbookSupport.openWorkbook(
            request.source(), request.formulaEnvironment(), bindings.workingDirectory())) {
      Objects.requireNonNull(workbook, "workbook must not be null");
      return Optional.empty();
    } catch (Exception exception) {
      return Optional.of(GridGrindProblems.fromException(exception, context));
    }
  }

  private static dev.erst.gridgrind.contract.dto.ProblemContext.OpenWorkbook openWorkbookContext(
      WorkbookPlan request, ExecutionInputBindings bindings) {
    return new dev.erst.gridgrind.contract.dto.ProblemContext.OpenWorkbook(
        ExecutionRequestPaths.requestShape(request),
        ExecutionRequestPaths.workbookReference(request, bindings.workingDirectory()));
  }

  private static ExecutionWorkbookSupport defaultWorkbookSupport() {
    return new ExecutionWorkbookSupport(GridGrindRequestDoctor::createTempWorkbookFile);
  }

  static Path createTempWorkbookFile(String prefix, String suffix) throws IOException {
    return Files.createTempFile(prefix, suffix);
  }
}
