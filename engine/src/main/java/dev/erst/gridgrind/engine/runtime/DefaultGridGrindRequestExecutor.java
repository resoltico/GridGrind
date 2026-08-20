package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.GridGrindProblemDetail;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.RequestWarning;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Default request executor that applies one GridGrind workflow against the workbook core. */
public final class DefaultGridGrindRequestExecutor implements GridGrindRequestExecutor {
  private final DefaultGridGrindRequestExecutorDependencies dependencies;
  private final StaticRequestValidator staticValidator;
  private final ExecutionResponseSupport responseSupport;

  /** Creates the production request executor with the default workbook executors and closers. */
  public DefaultGridGrindRequestExecutor() {
    this(DefaultGridGrindRequestExecutorDependencies.production());
  }

  /** Creates one executor from an explicit owned dependency bundle. */
  DefaultGridGrindRequestExecutor(DefaultGridGrindRequestExecutorDependencies dependencies) {
    this.dependencies = Objects.requireNonNull(dependencies, "dependencies must not be null");
    this.staticValidator = new StaticRequestValidator();
    this.responseSupport =
        new ExecutionResponseSupport(
            this.dependencies.workbookCloser(), this.dependencies.readableWorkbookCloser());
  }

  /** Executes one complete GridGrind request with optional live verbose journal emission. */
  @Override
  public WorkbookResult execute(
      WorkbookPlan request, ExecutionInputBindings bindings, ExecutionProgressSink sink) {
    return execute(request, bindings, sink, Optional.empty());
  }

  /** Executes one request while retaining raw-request locations for preflight diagnostics. */
  public WorkbookResult execute(
      WorkbookPlan request,
      ExecutionInputBindings bindings,
      ExecutionProgressSink sink,
      Optional<dev.erst.gridgrind.contract.json.RequestAnalysis> analysis) {
    WorkbookPlan authoredRequest = Objects.requireNonNull(request, "request must not be null");
    ExecutionInputBindings executionBindings =
        Objects.requireNonNull(bindings, "bindings must not be null");
    Objects.requireNonNull(analysis, "analysis must not be null");
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
        staticValidator.validate(authoredRequest).stream().findFirst();
    if (validationError.isPresent()) {
      validationPhase.fail(validationError.get().code());
      return ExecutionResponseSupport.failureResponse(
          protocolVersion,
          journal,
          authoredRequest,
          CalculationPolicyExecutor.notRequestedReport(authoredRequest.calculationPolicy()),
          validationError.get(),
          null,
          null);
    }
    validationPhase.succeed();

    ExecutionJournalRecorder.PhaseHandle inputResolutionPhase = journal.beginInputResolution();
    RequestPreflight.Result preflight =
        analysis
            .map(value -> RequestPreflight.verify(authoredRequest, executionBindings, value))
            .orElseGet(() -> RequestPreflight.verify(authoredRequest, executionBindings));
    try {
      if (!preflight.problems().isEmpty()) {
        GridGrindProblemDetail.Problem problem = preflight.problems().getFirst();
        inputResolutionPhase.fail(problem.code());
        return ExecutionResponseSupport.failureResponse(
            protocolVersion,
            journal,
            authoredRequest,
            CalculationPolicyExecutor.notRequestedReport(authoredRequest.calculationPolicy()),
            problem,
            null,
            null);
      }
      WorkbookPlan resolvedRequest = preflight.preparedPlan();
      ExecutionInputBindings preparedBindings = preflight.preparedBindings();
      inputResolutionPhase.succeed();

      List<RequestWarning> warnings =
          new java.util.ArrayList<>(GridGrindRequestWarnings.collect(resolvedRequest));
      warnings.addAll(preflight.warnings());

      ExecutionModeInput executionMode = executionMode(resolvedRequest);
      if (directEventReadEligible(resolvedRequest, executionMode)) {
        return responseSupport.guardUnexpectedRuntime(
            protocolVersion,
            resolvedRequest,
            journal,
            () ->
                workflowSupport.executeDirectEventReadWorkflow(
                    protocolVersion, resolvedRequest, warnings, journal, preparedBindings));
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
                    preparedBindings));
      }

      ExecutionJournalRecorder.PhaseHandle openPhase = journal.beginOpen();
      ExcelWorkbook workbook;
      try {
        workbook =
            workbookSupport.openWorkbook(
                resolvedRequest.source(), resolvedRequest.formulaEnvironment(), preparedBindings);
      } catch (Exception exception) {
        GridGrindProblemDetail.Problem problem =
            ExecutionResponseSupport.problemFor(
                exception,
                new dev.erst.gridgrind.contract.dto.ProblemContext.OpenWorkbook(
                    ExecutionRequestPaths.requestShape(resolvedRequest),
                    ExecutionRequestPaths.workbookReference(
                        resolvedRequest, preparedBindings.workingDirectory())));
        openPhase.fail(problem.code());
        return ExecutionResponseSupport.failureResponse(
            protocolVersion,
            journal,
            resolvedRequest,
            CalculationPolicyExecutor.notRequestedReport(resolvedRequest.calculationPolicy()),
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
                  preparedBindings));
    } finally {
      preflight.release();
    }
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

  static boolean directEventReadEligible(WorkbookPlan request, ExecutionModeInput executionMode) {
    return ExecutionWorkflowRouting.directEventReadEligible(request, executionMode);
  }

  static ExecutionModeInput executionMode(WorkbookPlan request) {
    return ExecutionWorkflowRouting.executionMode(request);
  }
}
