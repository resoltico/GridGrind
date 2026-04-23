package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.RequestWarning;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.excel.ExcelOoxmlPackageSecuritySupport;
import dev.erst.gridgrind.excel.ExcelStreamingWorkbookWriter;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookCommandExecutor;
import dev.erst.gridgrind.excel.WorkbookReadExecutor;
import java.nio.file.Files;
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
    this(
        new WorkbookCommandExecutor(),
        new WorkbookReadExecutor(),
        ExcelWorkbook::close,
        Files::createTempFile,
        ExcelOoxmlPackageSecuritySupport.ReadableWorkbook::close,
        ExcelStreamingWorkbookWriter::markRecalculateOnOpen);
  }

  /** Constructor for testing, allowing injection of custom workbook executors and closer. */
  DefaultGridGrindRequestExecutor(
      WorkbookCommandExecutor commandExecutor,
      WorkbookReadExecutor readExecutor,
      WorkbookCloser workbookCloser) {
    this(
        commandExecutor,
        readExecutor,
        workbookCloser,
        Files::createTempFile,
        ExcelOoxmlPackageSecuritySupport.ReadableWorkbook::close,
        ExcelStreamingWorkbookWriter::markRecalculateOnOpen);
  }

  /** Constructor for testing, allowing injection of custom executors, closer, and temp files. */
  DefaultGridGrindRequestExecutor(
      WorkbookCommandExecutor commandExecutor,
      WorkbookReadExecutor readExecutor,
      WorkbookCloser workbookCloser,
      TempFileFactory tempFileFactory) {
    this(
        commandExecutor,
        readExecutor,
        workbookCloser,
        tempFileFactory,
        ExcelOoxmlPackageSecuritySupport.ReadableWorkbook::close,
        ExcelStreamingWorkbookWriter::markRecalculateOnOpen);
  }

  /** Constructor for testing, allowing injection of direct-event readable-workbook closing. */
  DefaultGridGrindRequestExecutor(
      WorkbookCommandExecutor commandExecutor,
      WorkbookReadExecutor readExecutor,
      WorkbookCloser workbookCloser,
      TempFileFactory tempFileFactory,
      ReadableWorkbookCloser readableWorkbookCloser) {
    this(
        commandExecutor,
        readExecutor,
        workbookCloser,
        tempFileFactory,
        readableWorkbookCloser,
        ExcelStreamingWorkbookWriter::markRecalculateOnOpen);
  }

  /** Constructor for testing, allowing injection of direct-event closing and streaming calc. */
  DefaultGridGrindRequestExecutor(
      WorkbookCommandExecutor commandExecutor,
      WorkbookReadExecutor readExecutor,
      WorkbookCloser workbookCloser,
      TempFileFactory tempFileFactory,
      ReadableWorkbookCloser readableWorkbookCloser,
      StreamingCalculationApplier streamingCalculationApplier) {
    Objects.requireNonNull(commandExecutor, "commandExecutor must not be null");
    Objects.requireNonNull(readExecutor, "readExecutor must not be null");
    Objects.requireNonNull(workbookCloser, "workbookCloser must not be null");
    Objects.requireNonNull(tempFileFactory, "tempFileFactory must not be null");
    Objects.requireNonNull(readableWorkbookCloser, "readableWorkbookCloser must not be null");
    Objects.requireNonNull(
        streamingCalculationApplier, "streamingCalculationApplier must not be null");

    SemanticSelectorResolver selectorResolver = new SemanticSelectorResolver(readExecutor);
    AssertionExecutor assertionExecutor = new AssertionExecutor(readExecutor, selectorResolver);
    this.validationSupport = new ExecutionValidationSupport();
    this.workbookSupport = new ExecutionWorkbookSupport(tempFileFactory);
    this.stepSupport =
        new ExecutionStepSupport(
            commandExecutor, readExecutor, selectorResolver, assertionExecutor, tempFileFactory);
    this.responseSupport = new ExecutionResponseSupport(workbookCloser, readableWorkbookCloser);
    this.workflowSupport =
        new ExecutionWorkflowSupport(
            this.workbookSupport,
            new ExecutionCalculationSupport(streamingCalculationApplier),
            this.stepSupport,
            this.responseSupport,
            tempFileFactory);
  }

  /** Executes one complete GridGrind request with optional live verbose journal emission. */
  @Override
  public GridGrindResponse execute(
      WorkbookPlan request, ExecutionInputBindings bindings, ExecutionJournalSink sink) {
    ExecutionInputBindings executionBindings =
        bindings == null ? ExecutionInputBindings.processDefault() : bindings;
    ExecutionJournalRecorder journal =
        ExecutionJournalRecorder.start(request, sink, executionBindings.workingDirectory());
    GridGrindProtocolVersion protocolVersion =
        request == null ? GridGrindProtocolVersion.current() : request.protocolVersion();

    ExecutionJournalRecorder.PhaseHandle validationPhase = journal.beginValidation();
    if (request == null) {
      GridGrindResponse.Problem problem =
          GridGrindProblems.problem(
              GridGrindProblemCode.INVALID_REQUEST,
              "request must not be null",
              new GridGrindResponse.ProblemContext.ValidateRequest(null, null),
              (Throwable) null);
      validationPhase.fail("failed (" + problem.code() + ")");
      return ExecutionResponseSupport.failureResponse(
          protocolVersion, journal, 0, problem, null, null);
    }

    Optional<GridGrindResponse.Problem> validationError =
        validationSupport.validateRequest(request);
    if (validationError.isPresent()) {
      validationPhase.fail("failed (" + validationError.get().code() + ")");
      return ExecutionResponseSupport.failureResponse(
          protocolVersion,
          journal,
          request.steps().size(),
          CalculationPolicyExecutor.notRequestedReport(request.calculationPolicy()),
          validationError.get(),
          null,
          null);
    }
    validationPhase.succeed();

    ExecutionJournalRecorder.PhaseHandle inputResolutionPhase = journal.beginInputResolution();
    WorkbookPlan resolvedRequest;
    try {
      resolvedRequest = SourceBackedPlanResolver.resolve(request, executionBindings);
    } catch (Exception exception) {
      GridGrindResponse.Problem problem =
          ExecutionResponseSupport.problemFor(
              exception, stepSupport.resolveInputsContext(request, exception));
      inputResolutionPhase.fail("failed (" + problem.code() + ")");
      return ExecutionResponseSupport.failureResponse(
          protocolVersion,
          journal,
          request.steps().size(),
          CalculationPolicyExecutor.notRequestedReport(request.calculationPolicy()),
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
              request.source(), request.formulaEnvironment(), executionBindings.workingDirectory());
    } catch (Exception exception) {
      GridGrindResponse.Problem problem =
          ExecutionResponseSupport.problemFor(
              exception,
              new GridGrindResponse.ProblemContext.OpenWorkbook(
                  ExecutionRequestPaths.reqSourceType(request),
                  ExecutionRequestPaths.reqPersistenceType(request),
                  ExecutionRequestPaths.reqSourcePath(
                      request, executionBindings.workingDirectory())));
      openPhase.fail("failed (" + problem.code() + ")");
      return ExecutionResponseSupport.failureResponse(
          protocolVersion,
          journal,
          request.steps().size(),
          CalculationPolicyExecutor.notRequestedReport(request.calculationPolicy()),
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
