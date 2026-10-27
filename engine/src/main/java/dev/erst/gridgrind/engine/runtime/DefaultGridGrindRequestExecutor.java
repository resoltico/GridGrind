package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.plan.ExecutionModeSpec;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.GridGrindProblemDetail;
import dev.erst.gridgrind.core.problem.RequestWarning;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.core.vocabulary.ProtocolVersion;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbook;
import java.util.List;
import java.util.Objects;
import java.util.Optional;

/** Default request executor that applies one GridGrind workflow against the workbook core. */
public final class DefaultGridGrindRequestExecutor implements GridGrindRequestExecutor {
  private final DefaultGridGrindRequestExecutorDependencies dependencies;
  private final StaticRequestValidator staticValidator;
  private final ExecutionResults responseSupport;

  /** Creates the production request executor with the default workbook executors and closers. */
  public DefaultGridGrindRequestExecutor() {
    this(DefaultGridGrindRequestExecutorDependencies.production());
  }

  /** Creates one executor from an explicit owned dependency bundle. */
  DefaultGridGrindRequestExecutor(DefaultGridGrindRequestExecutorDependencies dependencies) {
    this.dependencies = Objects.requireNonNull(dependencies, "dependencies must not be null");
    this.staticValidator = new StaticRequestValidator();
    this.responseSupport =
        new ExecutionResults(
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
      Optional<dev.erst.gridgrind.core.problem.RequestIntake> analysis) {
    WorkbookPlan authoredRequest = Objects.requireNonNull(request, "request must not be null");
    ExecutionInputBindings executionBindings =
        Objects.requireNonNull(bindings, "bindings must not be null");
    Objects.requireNonNull(analysis, "analysis must not be null");
    TempFileFactory tempFileFactory = executionBindings.tempFileFactory();
    WorkbookLifecycle workbookSupport = new WorkbookLifecycle(tempFileFactory);
    StepExecution stepSupport = stepSupport(this.dependencies, tempFileFactory, executionBindings);
    CalculationWorkflow calculationSupport =
        new CalculationWorkflow(this.dependencies.streamingCalculationApplier());
    FullXssfWorkflow fullXssfWorkflow =
        new FullXssfWorkflow(
            calculationSupport,
            new FullXssfStepExecution(stepSupport, responseSupport),
            new FullXssfCalculationCheckpoint(calculationSupport, responseSupport),
            new FullXssfPersistence(workbookSupport, responseSupport),
            new HostAcceptanceExecutor(stepSupport),
            responseSupport);
    ExecutionDirectEventReadWorkflow directEventReadWorkflow =
        new ExecutionDirectEventReadWorkflow(stepSupport, responseSupport, tempFileFactory);
    ExecutionStreamingWorkflow streamingWorkflow =
        new ExecutionStreamingWorkflow(
            workbookSupport, calculationSupport, stepSupport, tempFileFactory);
    ExecutionJournalRecorder journal =
        ExecutionJournalRecorder.start(authoredRequest, sink, executionBindings.workingDirectory());
    ProtocolVersion protocolVersion = authoredRequest.protocolVersion();

    ExecutionJournalRecorder.PhaseHandle validationPhase = journal.beginValidation();
    Optional<GridGrindProblemDetail.Problem> validationError =
        staticValidator.validate(authoredRequest).stream().findFirst();
    if (validationError.isPresent()) {
      validationPhase.fail(validationError.get().code());
      return ExecutionResults.failureResponse(
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
        return ExecutionResults.failureResponse(
            protocolVersion,
            journal,
            authoredRequest,
            CalculationPolicyExecutor.notRequestedReport(authoredRequest.calculationPolicy()),
            problem,
            null,
            null);
      }
      var resolvedPlan = preflight.preparedPlan();
      WorkbookPlan resolvedRequest = resolvedPlan.request();
      ExecutionInputBindings preparedBindings = preflight.preparedBindings();
      inputResolutionPhase.succeed();

      List<RequestWarning> warnings =
          new java.util.ArrayList<>(GridGrindRequestWarnings.collect(resolvedRequest));
      warnings.addAll(preflight.warnings());

      ExecutionModeSpec executionMode = resolvedRequest.effectiveExecutionMode();
      if (directEventReadWorkflow.capabilities().accepts(resolvedRequest)) {
        return responseSupport.guardUnexpectedRuntime(
            protocolVersion,
            resolvedRequest,
            journal,
            () ->
                directEventReadWorkflow.execute(
                    new ExecutionDirectEventReadWorkflow.Call(
                        protocolVersion, resolvedPlan, warnings, journal, preparedBindings)));
      }
      if (streamingWorkflow.capabilities().accepts(resolvedRequest)) {
        return responseSupport.guardUnexpectedRuntime(
            protocolVersion,
            resolvedRequest,
            journal,
            () ->
                streamingWorkflow.execute(
                    new ExecutionStreamingWorkflow.Call(
                        protocolVersion,
                        resolvedPlan,
                        executionMode,
                        warnings,
                        journal,
                        preparedBindings)));
      }

      ExecutionJournalRecorder.PhaseHandle openPhase = journal.beginOpen();
      XssfWorkbook workbook;
      try {
        workbook =
            workbookSupport.openWorkbook(
                resolvedRequest.source(), resolvedRequest.formulaEnvironment(), preparedBindings);
      } catch (Exception exception) {
        GridGrindProblemDetail.Problem problem =
            GridGrindProblems.fromException(
                exception,
                new dev.erst.gridgrind.core.problem.ProblemContext.OpenWorkbook(
                    ExecutionRequestPaths.requestShape(resolvedRequest),
                    ExecutionRequestPaths.workbookReference(
                        resolvedRequest, preparedBindings.workingDirectory())));
        openPhase.fail(problem.code());
        return ExecutionResults.failureResponse(
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
              fullXssfWorkflow.execute(
                  new FullXssfWorkflow.Call(
                      protocolVersion,
                      resolvedPlan,
                      workbook,
                      executionMode,
                      warnings,
                      journal,
                      preparedBindings)));
    } finally {
      preflight.release();
    }
  }

  private static StepExecution stepSupport(
      DefaultGridGrindRequestExecutorDependencies dependencies,
      TempFileFactory tempFileFactory,
      ExecutionInputBindings bindings) {
    SemanticSelectorResolver selectorResolver =
        new SemanticSelectorResolver(dependencies.workbookEngine());
    AssertionExecutor assertionExecutor =
        new AssertionExecutor(dependencies.workbookEngine(), selectorResolver);
    return new StepExecution(
        dependencies.workbookEngine(),
        selectorResolver,
        assertionExecutor,
        tempFileFactory,
        bindings);
  }
}
