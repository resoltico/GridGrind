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
  private final NativeInvocationContext nativeContext;
  private final ExecutionResults responseSupport;

  /** Creates the production request executor with the default workbook executors and closers. */
  public DefaultGridGrindRequestExecutor(NativeInvocationContext nativeContext) {
    this(DefaultGridGrindRequestExecutorDependencies.production(), nativeContext);
  }

  /** Creates one executor from an explicit owned dependency bundle. */
  DefaultGridGrindRequestExecutor(
      DefaultGridGrindRequestExecutorDependencies dependencies,
      NativeInvocationContext nativeContext) {
    this.nativeContext = Objects.requireNonNull(nativeContext);
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
    nativeContext.requireOpen();
    ExecutionInputBindings runBindings = Objects.requireNonNull(bindings).forRun();
    ExecutionJournalRecorder journal =
        ExecutionJournalRecorder.start(request, sink, runBindings.workingDirectory());
    return dev.erst.gridgrind.engine.backend.xssf.WorkbookTempScope.call(
        runBindings.tempFileFactory()::createTempFile,
        () -> {
          WorkbookResult result;
          try {
            result = executeRun(request, runBindings, analysis, journal);
            result = withEstablishedFacts(result, request, runBindings.session());
          } catch (RuntimeException | Error failure) {
            try {
              runBindings.closeRunResources();
            } catch (java.io.IOException | RuntimeException | Error cleanup) {
              NativeInvocationContext.suppress(failure, cleanup);
            }
            nativeContext.closePreserving(failure);
            throw failure;
          }
          try {
            runBindings.closeRunResources();
          } catch (java.io.IOException cleanup) {
            var problem =
                GridGrindProblems.fromException(
                    cleanup,
                    new dev.erst.gridgrind.core.problem.ProblemContext.ExecuteRequest(
                        ExecutionRequestPaths.requestShape(request)));
            result =
                ExecutionResponseCloseFinalizer.withCloseFailure(
                    result, request, journal, null, null, null, problem);
          }
          try {
            nativeContext.close();
          } catch (RuntimeException cleanup) {
            var problem =
                GridGrindProblems.problem(
                    dev.erst.gridgrind.core.problem.GridGrindProblemCode.INTERNAL_ERROR,
                    "Native execution context cleanup failed.",
                    new dev.erst.gridgrind.core.problem.ProblemContext.ExecuteRequest(
                        ExecutionRequestPaths.requestShape(request)),
                    cleanup);
            result =
                dev.erst.gridgrind.core.result.WorkbookResults.withSecondaryFailure(
                    result, problem);
          }
          var finalizedJournal = journal.completePlan(result, request.steps().size());
          result = withJournal(result, finalizedJournal);
          return withRunDiagnostics(result, journal.diagnostics());
        });
  }

  private WorkbookResult executeRun(
      WorkbookPlan request,
      ExecutionInputBindings bindings,
      Optional<dev.erst.gridgrind.core.problem.RequestIntake> analysis,
      ExecutionJournalRecorder journal) {
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
        new ExecutionDirectEventReadWorkflow(stepSupport, responseSupport);
    ExecutionStreamingWorkflow streamingWorkflow =
        new ExecutionStreamingWorkflow(
            workbookSupport, calculationSupport, stepSupport, tempFileFactory);
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

    List<RequestWarning> warnings = executionBindings.session().warnings;
    warnings.addAll(GridGrindRequestWarnings.collect(resolvedRequest));
    warnings.addAll(preflight.warnings());

    ExecutionModeSpec executionMode = resolvedRequest.effectiveExecutionMode();
    var strategy = preflight.requirePreparedRequest().source().strategy();
    if (strategy.backend()
        == dev.erst.gridgrind.core.result.WorkbookExecutionEvidence.Backend.EVENT_READ) {
      return responseSupport.guardUnexpectedRuntime(
          protocolVersion,
          resolvedRequest,
          journal,
          () ->
              directEventReadWorkflow.execute(
                  new ExecutionDirectEventReadWorkflow.Call(
                      protocolVersion, preflight.requirePreparedRequest(), warnings, journal)));
    }
    if (strategy.backend()
        == dev.erst.gridgrind.core.result.WorkbookExecutionEvidence.Backend.STREAMING_WRITE) {
      return responseSupport.guardUnexpectedRuntime(
          protocolVersion,
          resolvedRequest,
          journal,
          () ->
              streamingWorkflow.execute(
                  new ExecutionStreamingWorkflow.Call(
                      protocolVersion,
                      preflight.requirePreparedRequest(),
                      executionMode,
                      warnings,
                      journal)));
    }

    ExecutionJournalRecorder.PhaseHandle openPhase = journal.beginOpen();
    XssfWorkbook workbook;
    try {
      var retained = preflight.requirePreparedRequest().source().takeXssf();
      workbook =
          retained.isPresent()
              ? retained.orElseThrow()
              : workbookSupport.openWorkbook(
                  resolvedRequest.source(), resolvedRequest.formulaEnvironment(), preparedBindings);
    } catch (Exception exception) {
      GridGrindProblemDetail.Problem problem =
          GridGrindProblems.fromException(
              exception,
              new dev.erst.gridgrind.core.problem.ProblemContext.OpenWorkbook(
                  ExecutionRequestPaths.requestShape(resolvedRequest),
                  ExecutionRequestPaths.workbookReference(
                      resolvedRequest, preparedBindings.workingDirectory()),
                  java.util.Optional.empty()));
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
                    preflight.requirePreparedRequest(),
                    workbook,
                    executionMode,
                    warnings,
                    journal)));
  }

  private static WorkbookResult withEstablishedFacts(
      WorkbookResult result, WorkbookPlan request, ExecutionSession session) {
    var facts = session.publication;
    var calculation = session.calculation.orElse(result.calculation());
    var persistence =
        facts.outcome()
                instanceof
                dev.erst.gridgrind.core.result.WorkbookResultPersistence.PublicationOutcome
                    .NotAttempted
            ? result.persistence()
            : dev.erst.gridgrind.core.result.WorkbookResults.publicationPersistenceOutcome(
                request, facts.outcome());
    var assertions =
        merge(
            session.planAssertions,
            result instanceof WorkbookResult.Success value
                ? value.assertions()
                : ((WorkbookResult.Failure) result).assertions());
    var inspections =
        merge(
            session.inspections,
            result instanceof WorkbookResult.Success value
                ? value.inspections()
                : ((WorkbookResult.Failure) result).inspections());
    var warnings =
        merge(
            session.warnings,
            result instanceof WorkbookResult.Success value
                ? value.warnings()
                : ((WorkbookResult.Failure) result).warnings());
    var evidence =
        session.prepared.isEmpty()
            ? (result instanceof WorkbookResult.Success value
                ? value.evidence()
                : ((WorkbookResult.Failure) result).evidence())
            : ExecutionEvidence.forExecution(
                request,
                session.prepared.orElseThrow().bindings(),
                calculation,
                assertions,
                session.hostAssertions,
                session.preservation,
                facts.stagedVerified);
    return switch (result) {
      case WorkbookResult.Success value ->
          new WorkbookResult.Success(
              value.protocolVersion(),
              value.planId(),
              value.journal(),
              calculation,
              persistence,
              evidence,
              warnings,
              assertions,
              inspections);
      case WorkbookResult.Failure value ->
          new WorkbookResult.Failure(
              value.protocolVersion(),
              value.planId(),
              value.journal(),
              calculation,
              persistence,
              evidence,
              warnings,
              assertions,
              inspections,
              value.problem());
    };
  }

  private static <T> List<T> merge(List<T> established, List<T> returned) {
    if (established.equals(returned)) return List.copyOf(established);
    var values = new java.util.ArrayList<T>(established);
    for (T value : returned) if (!values.contains(value)) values.add(value);
    return List.copyOf(values);
  }

  private static WorkbookResult withJournal(
      WorkbookResult result, dev.erst.gridgrind.core.result.ExecutionJournal journal) {
    return switch (result) {
      case WorkbookResult.Success value ->
          new WorkbookResult.Success(
              value.protocolVersion(),
              value.planId(),
              journal,
              value.calculation(),
              value.persistence(),
              value.evidence(),
              value.warnings(),
              value.assertions(),
              value.inspections());
      case WorkbookResult.Failure value ->
          new WorkbookResult.Failure(
              value.protocolVersion(),
              value.planId(),
              journal,
              value.calculation(),
              value.persistence(),
              value.evidence(),
              value.warnings(),
              value.assertions(),
              value.inspections(),
              value.problem());
    };
  }

  private static WorkbookResult withRunDiagnostics(
      WorkbookResult result, List<RequestWarning> diagnostics) {
    if (diagnostics.isEmpty()) {
      return result;
    }
    return switch (result) {
      case WorkbookResult.Success value -> {
        var warnings = new java.util.ArrayList<>(value.warnings());
        warnings.addAll(diagnostics);
        yield new WorkbookResult.Success(
            value.protocolVersion(),
            value.planId(),
            value.journal(),
            value.calculation(),
            value.persistence(),
            value.evidence(),
            warnings,
            value.assertions(),
            value.inspections());
      }
      case WorkbookResult.Failure value -> {
        var warnings = new java.util.ArrayList<>(value.warnings());
        warnings.addAll(diagnostics);
        yield new WorkbookResult.Failure(
            value.protocolVersion(),
            value.planId(),
            value.journal(),
            value.calculation(),
            value.persistence(),
            value.evidence(),
            warnings,
            value.assertions(),
            value.inspections(),
            value.problem());
      }
    };
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
