package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.RequestWarning;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Default request executor that applies one GridGrind workflow against the workbook core. */
public final class DefaultGridGrindRequestExecutor implements GridGrindRequestExecutor {
  private final DefaultGridGrindRequestExecutorDependencies dependencies;
  private final ExecutionValidationSupport validationSupport;
  private final ExecutionResponseSupport responseSupport;

  /** Creates the production request executor with the default workbook executors and closers. */
  public DefaultGridGrindRequestExecutor() {
    this(DefaultGridGrindRequestExecutorDependencies.production());
  }

  /** Creates one executor from an explicit owned dependency bundle. */
  DefaultGridGrindRequestExecutor(DefaultGridGrindRequestExecutorDependencies dependencies) {
    this.dependencies = Objects.requireNonNull(dependencies, "dependencies must not be null");
    this.validationSupport = new ExecutionValidationSupport();
    this.responseSupport =
        new ExecutionResponseSupport(
            this.dependencies.workbookCloser(), this.dependencies.readableWorkbookCloser());
  }

  /** Executes one complete GridGrind request with optional live verbose journal emission. */
  @Override
  public GridGrindResponse execute(
      WorkbookPlan request, ExecutionInputBindings bindings, ExecutionJournalSink sink) {
    WorkbookPlan authoredRequest = Objects.requireNonNull(request, "request must not be null");
    ExecutionInputBindings executionBindings =
        Objects.requireNonNull(bindings, "bindings must not be null");
    TempFileFactory tempFileFactory = executionBindings.tempFileFactory();
    ExecutionWorkbookSupport workbookSupport = new ExecutionWorkbookSupport(tempFileFactory);
    ExecutionStepSupport stepSupport = stepSupport(this.dependencies, tempFileFactory);
    ExecutionWorkflowSupport workflowSupport =
        new ExecutionWorkflowSupport(
            workbookSupport,
            new ExecutionCalculationSupport(this.dependencies.streamingCalculationApplier()),
            stepSupport,
            responseSupport,
            tempFileFactory);
    ExecutionJournalRecorder journal =
        ExecutionJournalRecorder.start(authoredRequest, sink, executionBindings.workingDirectory());
    GridGrindProtocolVersion protocolVersion = authoredRequest.protocolVersion();

    ExecutionJournalRecorder.PhaseHandle validationPhase = journal.beginValidation();
    Optional<GridGrindProblemDetail.Problem> validationError =
        validationSupport.firstValidationProblem(authoredRequest);
    if (validationError.isPresent()) {
      validationPhase.fail("failed (" + validationError.get().code() + ")");
      return ExecutionResponseSupport.failureResponse(
          protocolVersion,
          journal,
          authoredRequest.steps().size(),
          CalculationPolicyExecutor.notRequestedReport(authoredRequest.calculationPolicy()),
          validationError.get(),
          null,
          null);
    }
    validationPhase.succeed();

    ExecutionJournalRecorder.PhaseHandle inputResolutionPhase = journal.beginInputResolution();
    WorkbookPlan resolvedRequest;
    try {
      resolvedRequest = SourceBackedPlanResolver.resolve(authoredRequest, executionBindings);
    } catch (Exception exception) {
      GridGrindProblemDetail.Problem problem =
          ExecutionResponseSupport.problemFor(
              exception, stepSupport.resolveInputsContext(authoredRequest, exception));
      inputResolutionPhase.fail("failed (" + problem.code() + ")");
      return ExecutionResponseSupport.failureResponse(
          protocolVersion,
          journal,
          authoredRequest.steps().size(),
          CalculationPolicyExecutor.notRequestedReport(authoredRequest.calculationPolicy()),
          problem,
          null,
          null);
    }
    inputResolutionPhase.succeed();

    List<RequestWarning> warnings = GridGrindRequestWarnings.collect(resolvedRequest);

    ExecutionModeInput executionMode = executionMode(resolvedRequest);
    if (directEventReadEligible(resolvedRequest, executionMode)) {
      return responseSupport.guardUnexpectedRuntime(
          protocolVersion,
          resolvedRequest,
          journal,
          () ->
              workflowSupport.executeDirectEventReadWorkflow(
                  protocolVersion,
                  resolvedRequest,
                  warnings,
                  journal,
                  executionBindings.workingDirectory()));
    }
    if (executionMode instanceof ExecutionModeInput.StreamingWrite) {
      return responseSupport.guardUnexpectedRuntime(
          protocolVersion,
          resolvedRequest,
          journal,
          () ->
              workflowSupport.executeStreamingWorkflow(
                  protocolVersion,
                  resolvedRequest,
                  executionMode,
                  warnings,
                  journal,
                  executionBindings.workingDirectory()));
    }

    ExecutionJournalRecorder.PhaseHandle openPhase = journal.beginOpen();
    ExcelWorkbook workbook;
    try {
      workbook =
          workbookSupport.openWorkbook(
              authoredRequest.source(),
              authoredRequest.formulaEnvironment(),
              executionBindings.workingDirectory());
    } catch (Exception exception) {
      GridGrindProblemDetail.Problem problem =
          ExecutionResponseSupport.problemFor(
              exception,
              new dev.erst.gridgrind.contract.dto.ProblemContext.OpenWorkbook(
                  ExecutionRequestPaths.requestShape(authoredRequest),
                  ExecutionRequestPaths.workbookReference(
                      authoredRequest, executionBindings.workingDirectory())));
      openPhase.fail("failed (" + problem.code() + ")");
      return ExecutionResponseSupport.failureResponse(
          protocolVersion,
          journal,
          authoredRequest.steps().size(),
          CalculationPolicyExecutor.notRequestedReport(authoredRequest.calculationPolicy()),
          problem,
          null,
          null);
    }
    openPhase.succeed();

    return responseSupport.guardUnexpectedRuntime(
        protocolVersion,
        resolvedRequest,
        journal,
        workbook,
        () ->
            workflowSupport.executeWorkbookWorkflow(
                protocolVersion,
                resolvedRequest,
                workbook,
                executionMode,
                warnings,
                journal,
                executionBindings.workingDirectory()));
  }

  private static ExecutionStepSupport stepSupport(
      DefaultGridGrindRequestExecutorDependencies dependencies, TempFileFactory tempFileFactory) {
    SemanticSelectorResolver selectorResolver =
        new SemanticSelectorResolver(dependencies.workbookEngine());
    AssertionExecutor assertionExecutor =
        new AssertionExecutor(dependencies.workbookEngine(), selectorResolver);
    return new ExecutionStepSupport(
        dependencies.workbookEngine(), selectorResolver, assertionExecutor, tempFileFactory);
  }

  List<String> calculationPolicyFailures(WorkbookPlan request) {
    return ExecutionModeRules.calculationPolicyFailures(request);
  }

  List<String> executionModeFailures(WorkbookPlan request) { // LIM-019, LIM-020
    return ExecutionModeRules.executionModeFailures(request, executionMode(request));
  }

  static boolean directEventReadEligible(WorkbookPlan request, ExecutionModeInput executionMode) {
    return ExecutionModeRules.directEventReadEligible(request, executionMode);
  }

  static ExecutionModeInput executionMode(WorkbookPlan request) {
    return ExecutionModeRules.executionMode(request);
  }
}
