package dev.erst.gridgrind.executor;

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
  private final ExecutionValidationSupport validationSupport;
  private final ExecutionWorkbookSupport workbookSupport;
  private final ExecutionStepSupport stepSupport;
  private final ExecutionResponseSupport responseSupport;
  private final ExecutionWorkflowSupport workflowSupport;

  /** Creates the production request executor with the default workbook executors and closers. */
  public DefaultGridGrindRequestExecutor() {
    this(DefaultGridGrindRequestExecutorDependencies.production());
  }

  /** Creates one executor from an explicit owned dependency bundle. */
  DefaultGridGrindRequestExecutor(DefaultGridGrindRequestExecutorDependencies dependencies) {
    Objects.requireNonNull(dependencies, "dependencies must not be null");

    SemanticSelectorResolver selectorResolver =
        new SemanticSelectorResolver(dependencies.readExecutor());
    AssertionExecutor assertionExecutor =
        new AssertionExecutor(dependencies.readExecutor(), selectorResolver);
    this.validationSupport = new ExecutionValidationSupport();
    this.workbookSupport = new ExecutionWorkbookSupport(dependencies.tempFileFactory());
    this.stepSupport =
        new ExecutionStepSupport(
            dependencies.commandExecutor(),
            dependencies.readExecutor(),
            selectorResolver,
            assertionExecutor,
            dependencies.tempFileFactory());
    this.responseSupport =
        new ExecutionResponseSupport(
            dependencies.workbookCloser(), dependencies.readableWorkbookCloser());
    this.workflowSupport =
        new ExecutionWorkflowSupport(
            this.workbookSupport,
            new ExecutionCalculationSupport(dependencies.streamingCalculationApplier()),
            this.stepSupport,
            this.responseSupport,
            dependencies.tempFileFactory());
  }

  /** Executes one complete GridGrind request with optional live verbose journal emission. */
  @Override
  public GridGrindResponse execute(
      WorkbookPlan request, ExecutionInputBindings bindings, ExecutionJournalSink sink) {
    WorkbookPlan authoredRequest = Objects.requireNonNull(request, "request must not be null");
    ExecutionInputBindings executionBindings =
        Objects.requireNonNull(bindings, "bindings must not be null");
    ExecutionJournalRecorder journal =
        ExecutionJournalRecorder.start(authoredRequest, sink, executionBindings.workingDirectory());
    GridGrindProtocolVersion protocolVersion = authoredRequest.protocolVersion();

    ExecutionJournalRecorder.PhaseHandle validationPhase = journal.beginValidation();
    Optional<GridGrindProblemDetail.Problem> validationError =
        validationSupport.validateRequest(authoredRequest);
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
    journal.setWarnings(warnings);

    ExecutionModeSelection executionModes = executionModes(resolvedRequest);
    if (directEventReadEligible(resolvedRequest, executionModes)) {
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
    if (executionModes.writeMode() == ExecutionModeInput.WriteMode.STREAMING_WRITE) {
      return responseSupport.guardUnexpectedRuntime(
          protocolVersion,
          resolvedRequest,
          journal,
          () ->
              workflowSupport.executeStreamingWorkflow(
                  protocolVersion,
                  resolvedRequest,
                  executionModes,
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
                executionModes,
                warnings,
                journal,
                executionBindings.workingDirectory()));
  }

  Optional<String> calculationPolicyFailure(WorkbookPlan request) {
    return ExecutionModeRules.calculationPolicyFailure(request);
  }

  Optional<String> executionModeFailure(WorkbookPlan request) { // LIM-019, LIM-020
    return ExecutionModeRules.executionModeFailure(request, executionModes(request));
  }

  static boolean directEventReadEligible(
      WorkbookPlan request, ExecutionModeSelection executionModes) {
    return ExecutionModeRules.directEventReadEligible(request, executionModes);
  }

  static ExecutionModeSelection executionModes(WorkbookPlan request) {
    return ExecutionModeRules.executionModes(request);
  }
}
