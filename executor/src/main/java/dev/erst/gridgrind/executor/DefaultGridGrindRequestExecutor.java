package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.assertion.AssertionResult;
import dev.erst.gridgrind.contract.catalog.GridGrindContractText;
import dev.erst.gridgrind.contract.dto.CalculationExecutionStatus;
import dev.erst.gridgrind.contract.dto.CalculationPolicyInput;
import dev.erst.gridgrind.contract.dto.CalculationReport;
import dev.erst.gridgrind.contract.dto.CalculationStrategyInput;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.FormulaEnvironmentInput;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindProtocolVersion;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.NamedRangeScope;
import dev.erst.gridgrind.contract.dto.PivotTableInput;
import dev.erst.gridgrind.contract.dto.RequestWarning;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.ChartSelector;
import dev.erst.gridgrind.contract.selector.ColumnBandSelector;
import dev.erst.gridgrind.contract.selector.DrawingObjectSelector;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.RowBandSelector;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.TableCellSelector;
import dev.erst.gridgrind.contract.selector.TableRowSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.contract.step.WorkbookStep;
import dev.erst.gridgrind.excel.CellNotFoundException;
import dev.erst.gridgrind.excel.ExcelEventWorkbookReader;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.ExcelOoxmlPackageSecuritySnapshot;
import dev.erst.gridgrind.excel.ExcelOoxmlPackageSecuritySupport;
import dev.erst.gridgrind.excel.ExcelOoxmlPersistenceOptions;
import dev.erst.gridgrind.excel.ExcelStreamingWorkbookWriter;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.FormulaException;
import dev.erst.gridgrind.excel.InvalidCellAddressException;
import dev.erst.gridgrind.excel.InvalidRangeAddressException;
import dev.erst.gridgrind.excel.NamedRangeNotFoundException;
import dev.erst.gridgrind.excel.SheetNotFoundException;
import dev.erst.gridgrind.excel.WorkbookCommand;
import dev.erst.gridgrind.excel.WorkbookCommandExecutor;
import dev.erst.gridgrind.excel.WorkbookLocation;
import dev.erst.gridgrind.excel.WorkbookReadCommand;
import dev.erst.gridgrind.excel.WorkbookReadExecutor;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.Optional;
import java.util.Set;

/**
 * Default request executor that applies the GridGrind workflow against the workbook core.
 *
 * <p>This orchestration seam intentionally spans the full contract-plus-engine vocabulary.
 */
@SuppressWarnings("PMD.ExcessiveImports")
public final class DefaultGridGrindRequestExecutor implements GridGrindRequestExecutor {
  private static final Set<Class<? extends MutationAction>> STREAMING_WRITE_MUTATION_ACTION_TYPES =
      GridGrindContractText.streamingWriteMutationActionClasses();

  private static final Set<Class<? extends InspectionQuery>> EVENT_READ_INSPECTION_QUERY_TYPES =
      GridGrindContractText.eventReadInspectionQueryClasses();

  private final WorkbookCommandExecutor commandExecutor;
  private final WorkbookReadExecutor readExecutor;
  private final SemanticSelectorResolver selectorResolver;
  private final AssertionExecutor assertionExecutor;
  private final WorkbookCloser workbookCloser;
  private final TempFileFactory tempFileFactory;
  private final ReadableWorkbookCloser readableWorkbookCloser;
  private final StreamingCalculationApplier streamingCalculationApplier;

  /** Creates the production request executor with the default workbook executors and closer. */
  public DefaultGridGrindRequestExecutor() {
    this(
        new WorkbookCommandExecutor(),
        new WorkbookReadExecutor(),
        ExcelWorkbook::close,
        Files::createTempFile,
        ExcelOoxmlPackageSecuritySupport.ReadableWorkbook::close,
        ExcelStreamingWorkbookWriter::markRecalculateOnOpen);
  }

  /** Constructor for testing, allowing injection of custom executors and closer. */
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
    this.commandExecutor =
        Objects.requireNonNull(commandExecutor, "commandExecutor must not be null");
    this.readExecutor = Objects.requireNonNull(readExecutor, "readExecutor must not be null");
    this.selectorResolver = new SemanticSelectorResolver(this.readExecutor);
    this.assertionExecutor = new AssertionExecutor(this.readExecutor, this.selectorResolver);
    this.workbookCloser = Objects.requireNonNull(workbookCloser, "workbookCloser must not be null");
    this.tempFileFactory =
        Objects.requireNonNull(tempFileFactory, "tempFileFactory must not be null");
    this.readableWorkbookCloser =
        Objects.requireNonNull(readableWorkbookCloser, "readableWorkbookCloser must not be null");
    this.streamingCalculationApplier =
        Objects.requireNonNull(
            streamingCalculationApplier, "streamingCalculationApplier must not be null");
  }

  /** Executes one complete GridGrind request with optional live verbose journal emission. */
  @Override
  public GridGrindResponse execute(
      WorkbookPlan request, ExecutionInputBindings bindings, ExecutionJournalSink sink) {
    ExecutionInputBindings executionBindings =
        bindings == null ? ExecutionInputBindings.processDefault() : bindings;
    ExecutionJournalRecorder journal = ExecutionJournalRecorder.start(request, sink);
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
      return failureResponse(protocolVersion, journal, 0, problem, null, null);
    }

    Optional<GridGrindResponse.Problem> validationError = validateRequest(request);
    if (validationError.isPresent()) {
      validationPhase.fail("failed (" + validationError.get().code() + ")");
      return failureResponse(
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
          problemFor(exception, resolveInputsContext(request, exception));
      inputResolutionPhase.fail("failed (" + problem.code() + ")");
      return failureResponse(
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
      return guardUnexpectedRuntime(
          protocolVersion,
          resolvedRequest,
          journal,
          () ->
              executeDirectEventReadWorkflow(protocolVersion, resolvedRequest, warnings, journal));
    }
    if (executionModes.writeMode() == ExecutionModeInput.WriteMode.STREAMING_WRITE) {
      return guardUnexpectedRuntime(
          protocolVersion,
          resolvedRequest,
          journal,
          () ->
              executeStreamingWorkflow(
                  protocolVersion, resolvedRequest, executionModes, warnings, journal));
    }

    ExecutionJournalRecorder.PhaseHandle openPhase = journal.beginOpen();
    ExcelWorkbook workbook;
    try {
      workbook = openWorkbook(request.source(), request.formulaEnvironment());
    } catch (Exception exception) {
      GridGrindResponse.Problem problem =
          problemFor(
              exception,
              new GridGrindResponse.ProblemContext.OpenWorkbook(
                  reqSourceType(request), reqPersistenceType(request), reqSourcePath(request)));
      openPhase.fail("failed (" + problem.code() + ")");
      return failureResponse(
          protocolVersion,
          journal,
          request.steps().size(),
          CalculationPolicyExecutor.notRequestedReport(request.calculationPolicy()),
          problem,
          null,
          null);
    }
    openPhase.succeed();

    return guardUnexpectedRuntime(
        protocolVersion,
        resolvedRequest,
        journal,
        workbook,
        () ->
            executeWorkbookWorkflow(
                protocolVersion, resolvedRequest, workbook, executionModes, warnings, journal));
  }

  private GridGrindResponse executeWorkbookWorkflow(
      GridGrindProtocolVersion protocolVersion,
      WorkbookPlan request,
      ExcelWorkbook workbook,
      ExecutionModeSelection executionModes,
      List<RequestWarning> warnings,
      ExecutionJournalRecorder journal) {
    WorkbookLocation workbookLocation =
        workbookLocationFor(request.source(), request.persistence());
    List<AssertionResult> assertions = new ArrayList<>();
    List<InspectionResult> inspections = new ArrayList<>();
    CalculationReport calculation =
        CalculationPolicyExecutor.notRequestedReport(request.calculationPolicy());
    boolean calculationExecuted = false;
    for (int stepIndex = 0; stepIndex < request.steps().size(); stepIndex++) {
      WorkbookStep step = request.steps().get(stepIndex);
      if (!calculationExecuted && shouldExecuteCalculationBeforeStep(request, step)) {
        CalculationExecutionOutcome calculationOutcome =
            executeCalculationPolicy(workbook, request, journal);
        calculation = calculationOutcome.report();
        calculationExecuted = true;
        if (calculationOutcome.failure() != null) {
          return closeWorkbook(
              workbook,
              failureResponse(
                  protocolVersion,
                  journal,
                  request.steps().size(),
                  calculation,
                  calculationOutcome.failure(),
                  null,
                  null),
              request,
              journal,
              calculationOutcome.failure().code(),
              null,
              null);
        }
      }
      ExecutionJournalRecorder.StepHandle stepHandle = journal.beginStep(stepIndex, step);
      try {
        switch (step) {
          case MutationStep mutationStep -> executeMutationStep(workbook, mutationStep);
          case AssertionStep assertionStep ->
              assertions.add(
                  executeAssertionStep(
                      assertionStep, workbook, workbookLocation, executionModes.readMode()));
          case InspectionStep inspectionStep ->
              inspections.add(
                  executeInspectionStep(
                      inspectionStep, workbook, workbookLocation, executionModes.readMode()));
        }
        stepHandle.succeed();
      } catch (Exception exception) {
        GridGrindResponse.Problem problem =
            problemFor(exception, executeStepContext(request, stepIndex, step, exception));
        stepHandle.fail(
            problem.code(), problem.category(), problem.context().stage(), problem.message());
        return closeWorkbook(
            workbook,
            failureResponse(
                protocolVersion,
                journal,
                request.steps().size(),
                calculation,
                problem,
                stepIndex,
                step.stepId()),
            request,
            journal,
            problem.code(),
            stepIndex,
            step.stepId());
      }
    }

    if (!calculationExecuted) {
      CalculationExecutionOutcome calculationOutcome =
          executeCalculationPolicy(workbook, request, journal);
      calculation = calculationOutcome.report();
      if (calculationOutcome.failure() != null) {
        return closeWorkbook(
            workbook,
            failureResponse(
                protocolVersion,
                journal,
                request.steps().size(),
                calculation,
                calculationOutcome.failure(),
                null,
                null),
            request,
            journal,
            calculationOutcome.failure().code(),
            null,
            null);
      }
    }

    ExecutionJournalRecorder.PhaseHandle persistencePhase = journal.beginPersistence();
    GridGrindResponse.PersistenceOutcome persistence;
    try {
      persistence = persistWorkbook(workbook, request.source(), request.persistence());
    } catch (Exception exception) {
      GridGrindResponse.Problem problem =
          problemFor(
              exception,
              new GridGrindResponse.ProblemContext.PersistWorkbook(
                  reqSourceType(request),
                  reqPersistenceType(request),
                  reqSourcePath(request),
                  persistencePath(request.source(), request.persistence())));
      persistencePhase.fail("failed (" + problem.code() + ")");
      return closeWorkbook(
          workbook,
          failureResponse(
              protocolVersion, journal, request.steps().size(), calculation, problem, null, null),
          request,
          journal,
          problem.code(),
          null,
          null);
    }
    persistencePhase.succeed();

    return closeWorkbook(
        workbook,
        new GridGrindResponse.Success(
            protocolVersion,
            journal.buildSuccess(request.steps().size()),
            calculation,
            persistence,
            warnings,
            List.copyOf(assertions),
            List.copyOf(inspections)),
        request,
        journal,
        null,
        null,
        null);
  }

  /** Validates cross-field constraints that cannot be enforced at the record level. */
  private Optional<GridGrindResponse.Problem> validateRequest(WorkbookPlan request) {
    Optional<String> calculationPolicyFailure = calculationPolicyFailure(request);
    if (calculationPolicyFailure.isPresent()) {
      return Optional.of(
          GridGrindProblems.problem(
              GridGrindProblemCode.INVALID_REQUEST,
              calculationPolicyFailure.get(),
              new GridGrindResponse.ProblemContext.ValidateRequest(
                  reqSourceType(request), reqPersistenceType(request)),
              (Throwable) null));
    }
    Optional<String> executionModeFailure = executionModeFailure(request);
    if (executionModeFailure.isPresent()) {
      return Optional.of(
          GridGrindProblems.problem(
              GridGrindProblemCode.INVALID_REQUEST,
              executionModeFailure.get(),
              new GridGrindResponse.ProblemContext.ValidateRequest(
                  reqSourceType(request), reqPersistenceType(request)),
              (Throwable) null));
    }
    return switch (request.persistence()) {
      case WorkbookPlan.WorkbookPersistence.OverwriteSource _ ->
          switch (request.source()) {
            case WorkbookPlan.WorkbookSource.New _ ->
                Optional.of(
                    GridGrindProblems.problem(
                        GridGrindProblemCode.INVALID_REQUEST,
                        "OVERWRITE persistence requires an EXISTING source; "
                            + "a NEW workbook has no source file to overwrite",
                        new GridGrindResponse.ProblemContext.ValidateRequest(
                            reqSourceType(request), reqPersistenceType(request)),
                        (Throwable) null));
            case WorkbookPlan.WorkbookSource.ExistingFile _ -> Optional.empty();
          };
      case WorkbookPlan.WorkbookPersistence.None _ -> Optional.empty();
      case WorkbookPlan.WorkbookPersistence.SaveAs _ -> Optional.empty();
    };
  }

  private GridGrindResponse executeDirectEventReadWorkflow(
      GridGrindProtocolVersion protocolVersion,
      WorkbookPlan request,
      List<RequestWarning> warnings,
      ExecutionJournalRecorder journal) {
    CalculationReport calculation =
        CalculationPolicyExecutor.notRequestedReport(request.calculationPolicy());
    WorkbookPlan.WorkbookSource.ExistingFile source =
        (WorkbookPlan.WorkbookSource.ExistingFile) request.source();
    List<InspectionResult> inspections = new ArrayList<>();
    ExecutionJournalRecorder.PhaseHandle openPhase = journal.beginOpen();
    ExcelOoxmlPackageSecuritySupport.ReadableWorkbook materialized;
    try {
      materialized =
          ExcelOoxmlPackageSecuritySupport.materializeReadableWorkbook(
              Path.of(source.path()),
              OoxmlPackageSecurityConverter.toExcelOpenOptions(source.security()),
              tempFileFactory::createTempFile);
    } catch (Exception exception) {
      GridGrindResponse.Problem problem =
          problemFor(
              exception,
              new GridGrindResponse.ProblemContext.OpenWorkbook(
                  reqSourceType(request), reqPersistenceType(request), reqSourcePath(request)));
      openPhase.fail("failed (" + problem.code() + ")");
      return closeReadableWorkbook(
          null,
          failureResponse(
              protocolVersion, journal, request.steps().size(), calculation, problem, null, null),
          request,
          journal,
          problem.code(),
          null,
          null);
    }
    openPhase.succeed();
    for (int stepIndex = 0; stepIndex < request.steps().size(); stepIndex++) {
      InspectionStep inspectionStep = (InspectionStep) request.steps().get(stepIndex);
      ExecutionJournalRecorder.StepHandle stepHandle = journal.beginStep(stepIndex, inspectionStep);
      try {
        inspections.add(executeEventInspection(materialized.workbookPath(), inspectionStep));
        stepHandle.succeed();
      } catch (Exception exception) {
        GridGrindResponse.Problem problem =
            problemFor(
                exception, executeStepContext(request, stepIndex, inspectionStep, exception));
        stepHandle.fail(
            problem.code(), problem.category(), problem.context().stage(), problem.message());
        return closeReadableWorkbook(
            materialized,
            failureResponse(
                protocolVersion,
                journal,
                request.steps().size(),
                calculation,
                problem,
                stepIndex,
                inspectionStep.stepId()),
            request,
            journal,
            problem.code(),
            stepIndex,
            inspectionStep.stepId());
      }
    }
    return closeReadableWorkbook(
        materialized,
        new GridGrindResponse.Success(
            protocolVersion,
            journal.buildSuccess(request.steps().size()),
            calculation,
            new GridGrindResponse.PersistenceOutcome.NotSaved(),
            warnings,
            List.of(),
            List.copyOf(inspections)),
        request,
        journal,
        null,
        null,
        null);
  }

  private GridGrindResponse executeStreamingWorkflow(
      GridGrindProtocolVersion protocolVersion,
      WorkbookPlan request,
      ExecutionModeSelection executionModes,
      List<RequestWarning> warnings,
      ExecutionJournalRecorder journal) {
    WorkbookLocation workbookLocation =
        workbookLocationFor(request.source(), request.persistence());
    List<AssertionResult> assertions = new ArrayList<>();
    List<InspectionResult> inspections = new ArrayList<>();
    CalculationReport calculation =
        CalculationPolicyExecutor.notRequestedReport(request.calculationPolicy());
    Path materializedPath = null;
    boolean movedToPersistenceTarget = false;
    ExecutionJournalRecorder.PhaseHandle openPhase = journal.beginOpen();
    openPhase.succeed();
    try (ExcelStreamingWorkbookWriter writer = new ExcelStreamingWorkbookWriter()) {
      for (int stepIndex = 0; stepIndex < request.steps().size(); stepIndex++) {
        WorkbookStep step = request.steps().get(stepIndex);
        ExecutionJournalRecorder.StepHandle stepHandle = journal.beginStep(stepIndex, step);
        try {
          switch (step) {
            case MutationStep mutationStep -> executeStreamingMutationStep(writer, mutationStep);
            case AssertionStep assertionStep ->
                assertions.add(
                    executeStreamingAssertionStep(writer, assertionStep, workbookLocation));
            case InspectionStep inspectionStep ->
                inspections.add(
                    executeStreamingInspectionStep(
                        writer, inspectionStep, workbookLocation, executionModes.readMode()));
          }
          stepHandle.succeed();
        } catch (Exception exception) {
          deleteIfExists(materializedPath);
          GridGrindResponse.Problem problem =
              problemFor(exception, executeStepContext(request, stepIndex, step, exception));
          stepHandle.fail(
              problem.code(), problem.category(), problem.context().stage(), problem.message());
          return failureResponse(
              protocolVersion,
              journal,
              request.steps().size(),
              calculation,
              problem,
              stepIndex,
              step.stepId());
        }
      }

      CalculationExecutionOutcome calculationOutcome =
          executeStreamingCalculationPolicy(writer, request, journal);
      calculation = calculationOutcome.report();
      if (calculationOutcome.failure() != null) {
        deleteIfExists(materializedPath);
        return failureResponse(
            protocolVersion,
            journal,
            request.steps().size(),
            calculation,
            calculationOutcome.failure(),
            null,
            null);
      }

      materializedPath = tempFileFactory.createTempFile("gridgrind-streaming-write-", ".xlsx");
      writer.save(materializedPath);
    } catch (IOException exception) {
      deleteIfExists(materializedPath);
      GridGrindResponse.Problem problem =
          problemFor(
              exception,
              new GridGrindResponse.ProblemContext.ExecuteRequest(
                  reqSourceType(request), reqPersistenceType(request)));
      return failureResponse(
          protocolVersion, journal, request.steps().size(), calculation, problem, null, null);
    }

    ExecutionJournalRecorder.PhaseHandle persistencePhase = journal.beginPersistence();
    GridGrindResponse.PersistenceOutcome persistence;
    try {
      persistence =
          persistStreamingWorkbook(materializedPath, request.persistence(), request.source());
      movedToPersistenceTarget =
          !(persistence instanceof GridGrindResponse.PersistenceOutcome.NotSaved);
    } catch (Exception exception) {
      deleteIfExists(materializedPath);
      GridGrindResponse.Problem problem =
          problemFor(
              exception,
              new GridGrindResponse.ProblemContext.PersistWorkbook(
                  reqSourceType(request),
                  reqPersistenceType(request),
                  reqSourcePath(request),
                  persistencePath(request.source(), request.persistence())));
      persistencePhase.fail("failed (" + problem.code() + ")");
      return failureResponse(
          protocolVersion, journal, request.steps().size(), calculation, problem, null, null);
    } finally {
      if (!movedToPersistenceTarget) {
        deleteIfExists(materializedPath);
      }
    }
    persistencePhase.succeed();

    return new GridGrindResponse.Success(
        protocolVersion,
        journal.buildSuccess(request.steps().size()),
        calculation,
        persistence,
        warnings,
        List.copyOf(assertions),
        List.copyOf(inspections));
  }

  private void executeMutationStep(ExcelWorkbook workbook, MutationStep mutationStep)
      throws IOException {
    Selector resolvedTarget =
        selectorResolver.resolveMutationTarget(
            workbook, mutationStep.target(), mutationStep.action());
    WorkbookCommand command =
        WorkbookCommandConverter.toCommand(resolvedTarget, mutationStep.action());
    commandExecutor.apply(workbook, command);
  }

  private void executeStreamingMutationStep(
      ExcelStreamingWorkbookWriter writer, MutationStep mutationStep) throws IOException {
    WorkbookCommand command = WorkbookCommandConverter.toCommand(mutationStep);
    writer.apply(command);
  }

  private GridGrindResponse.ProblemContext.ExecuteStep executeStepContext(
      WorkbookPlan request, int stepIndex, WorkbookStep step, Exception exception) {
    return new GridGrindResponse.ProblemContext.ExecuteStep(
        reqSourceType(request),
        reqPersistenceType(request),
        stepIndex,
        step.stepId(),
        step.stepKind(),
        stepType(step),
        sheetNameFor(step, exception),
        addressFor(step, exception),
        rangeFor(step, exception),
        formulaFor(step, exception),
        namedRangeNameFor(step, exception));
  }

  private GridGrindResponse.ProblemContext resolveInputsContext(
      WorkbookPlan request, Exception exception) {
    return new GridGrindResponse.ProblemContext.ResolveInputs(
        reqSourceType(request),
        reqPersistenceType(request),
        exception instanceof InputSourceException inputSourceException
            ? inputSourceException.inputKind()
            : null,
        exception instanceof InputSourceException inputSourceException
            ? inputSourceException.inputPath()
            : null);
  }

  private InspectionResult executeInspectionStep(
      InspectionStep inspectionStep,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation,
      ExecutionModeInput.ReadMode readMode)
      throws IOException {
    return switch (readMode) {
      case ExecutionModeInput.ReadMode.FULL_XSSF ->
          executeFullInspectionStep(inspectionStep, workbook, workbookLocation);
      case ExecutionModeInput.ReadMode.EVENT_READ ->
          executeEventInspectionAgainstWorkbook(inspectionStep, workbook);
    };
  }

  AssertionResult executeAssertionStep(
      AssertionStep assertionStep,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation,
      ExecutionModeInput.ReadMode readMode)
      throws IOException, AssertionFailedException {
    return switch (readMode) {
      case ExecutionModeInput.ReadMode.FULL_XSSF ->
          assertionExecutor.execute(assertionStep, workbook, workbookLocation);
      case ExecutionModeInput.ReadMode.EVENT_READ ->
          throw new IllegalStateException(
              "executionMode.readMode=EVENT_READ does not support assertion steps");
    };
  }

  private InspectionResult executeFullInspectionStep(
      InspectionStep inspectionStep, ExcelWorkbook workbook, WorkbookLocation workbookLocation) {
    SemanticSelectorResolver.ResolvedInspectionTarget resolvedTarget =
        selectorResolver.resolveInspectionTarget(
            inspectionStep.stepId(), workbook, inspectionStep.target(), inspectionStep.query());
    if (resolvedTarget.isShortCircuit()) {
      return resolvedTarget.shortCircuitResult();
    }
    dev.erst.gridgrind.excel.WorkbookReadResult result =
        readExecutor
            .apply(
                workbook,
                workbookLocation,
                InspectionCommandConverter.toReadCommand(
                    inspectionStep.stepId(), resolvedTarget.selector(), inspectionStep.query()))
            .getFirst();
    return InspectionResultConverter.toReadResult(result);
  }

  private InspectionResult executeEventInspectionAgainstWorkbook(
      InspectionStep inspectionStep, ExcelWorkbook workbook) throws IOException {
    Path tempPath = null;
    try {
      tempPath = tempFileFactory.createTempFile("gridgrind-event-read-", ".xlsx");
      workbook.savePlainWorkbook(tempPath);
      return executeEventInspection(tempPath, inspectionStep);
    } finally {
      deleteIfExists(tempPath);
    }
  }

  private InspectionResult executeStreamingInspectionStep(
      ExcelStreamingWorkbookWriter writer,
      InspectionStep inspectionStep,
      WorkbookLocation workbookLocation,
      ExecutionModeInput.ReadMode readMode)
      throws IOException {
    Path tempPath = tempFileFactory.createTempFile("gridgrind-streaming-step-", ".xlsx");
    try {
      writer.save(tempPath);
      return executeInspectionAgainstMaterializedPath(
          inspectionStep, workbookLocation, readMode, tempPath);
    } finally {
      deleteIfExists(tempPath);
    }
  }

  private AssertionResult executeStreamingAssertionStep(
      ExcelStreamingWorkbookWriter writer,
      AssertionStep assertionStep,
      WorkbookLocation workbookLocation)
      throws IOException, AssertionFailedException {
    Path tempPath = tempFileFactory.createTempFile("gridgrind-streaming-step-", ".xlsx");
    try {
      writer.save(tempPath);
      return executeFullAssertionAgainstMaterializedPath(assertionStep, workbookLocation, tempPath);
    } finally {
      deleteIfExists(tempPath);
    }
  }

  InspectionResult executeInspectionAgainstMaterializedPath(
      InspectionStep inspectionStep,
      WorkbookLocation workbookLocation,
      ExecutionModeInput.ReadMode readMode,
      Path materializedPath)
      throws IOException {
    return switch (readMode) {
      case ExecutionModeInput.ReadMode.FULL_XSSF ->
          executeFullInspectionAgainstMaterializedPath(
              inspectionStep, workbookLocation, materializedPath);
      case ExecutionModeInput.ReadMode.EVENT_READ ->
          executeEventInspection(materializedPath, inspectionStep);
    };
  }

  private InspectionResult executeFullInspectionAgainstMaterializedPath(
      InspectionStep inspectionStep, WorkbookLocation workbookLocation, Path materializedPath)
      throws IOException {
    try (ExcelWorkbook workbook =
        ExcelWorkbook.open(
            materializedPath, FormulaEnvironmentConverter.toExcelFormulaEnvironment(null))) {
      return executeFullInspectionStep(inspectionStep, workbook, workbookLocation);
    }
  }

  private AssertionResult executeFullAssertionAgainstMaterializedPath(
      AssertionStep assertionStep, WorkbookLocation workbookLocation, Path materializedPath)
      throws IOException, AssertionFailedException {
    try (ExcelWorkbook workbook =
        ExcelWorkbook.open(
            materializedPath, FormulaEnvironmentConverter.toExcelFormulaEnvironment(null))) {
      return assertionExecutor.execute(assertionStep, workbook, workbookLocation);
    }
  }

  private InspectionResult executeEventInspection(Path workbookPath, InspectionStep inspectionStep)
      throws IOException {
    ExcelEventWorkbookReader eventWorkbookReader = new ExcelEventWorkbookReader();
    WorkbookReadCommand command = InspectionCommandConverter.toReadCommand(inspectionStep);
    dev.erst.gridgrind.excel.WorkbookReadResult result =
        eventWorkbookReader
            .apply(workbookPath, List.of((WorkbookReadCommand.Introspection) command))
            .getFirst();
    return InspectionResultConverter.toReadResult(result);
  }

  Optional<String> calculationPolicyFailure(WorkbookPlan request) {
    if (!CalculationPolicyExecutor.requiresMutationPrefix(request.calculationPolicy())) {
      return Optional.empty();
    }
    boolean seenObservationStep = false;
    for (WorkbookStep step : request.steps()) {
      if (step instanceof MutationStep) {
        if (seenObservationStep) {
          return Optional.of(
              "execution.calculation.strategy="
                  + request.calculationPolicy().effectiveStrategy().strategyType()
                  + " requires all MUTATION steps to appear before any ASSERTION or INSPECTION"
                  + " step so calculation can run once at the mutation-to-observation boundary");
        }
      } else {
        seenObservationStep = true;
      }
    }
    return Optional.empty();
  }

  Optional<String> executionModeFailure(WorkbookPlan request) { // LIM-019, LIM-020
    ExecutionModeSelection executionModes = executionModes(request);
    if (executionModes.readMode() == ExecutionModeInput.ReadMode.EVENT_READ) {
      if (!CalculationPolicyExecutor.allowsEventRead(request.calculationPolicy())) {
        return Optional.of(
            "execution.mode.readMode=EVENT_READ requires execution.calculation.strategy="
                + "DO_NOT_CALCULATE and markRecalculateOnOpen=false");
      }
      for (WorkbookStep step : request.steps()) {
        if (!(step instanceof InspectionStep inspectionStep)) {
          return Optional.of(
              "execution.mode.readMode=EVENT_READ supports inspection steps only; unsupported step"
                  + " kind: "
                  + step.stepKind());
        }
        if (!EVENT_READ_INSPECTION_QUERY_TYPES.contains(inspectionStep.query().getClass())) {
          return Optional.of(
              "execution.mode.readMode=EVENT_READ supports "
                  + GridGrindContractText.eventReadInspectionQueryTypePhrase()
                  + " only; unsupported inspection query type: "
                  + inspectionStep.query().queryType());
        }
      }
    }
    if (executionModes.writeMode() == ExecutionModeInput.WriteMode.STREAMING_WRITE) {
      if (!CalculationPolicyExecutor.allowsStreamingWrite(request.calculationPolicy())) {
        return Optional.of(
            "execution.mode.writeMode=STREAMING_WRITE requires"
                + " execution.calculation.strategy=DO_NOT_CALCULATE because low-memory streaming"
                + " writes do not support immediate server-side evaluation or cache clearing");
      }
      if (!(request.source() instanceof WorkbookPlan.WorkbookSource.New)) {
        return Optional.of(
            "execution.mode.writeMode=STREAMING_WRITE requires source.type=NEW because"
                + " low-memory streaming writes do not author in-place edits on EXISTING"
                + " workbooks");
      }
      boolean seenEnsureSheet = false;
      for (WorkbookStep step : request.steps()) {
        switch (step) {
          case MutationStep mutationStep -> {
            if (!STREAMING_WRITE_MUTATION_ACTION_TYPES.contains(mutationStep.action().getClass())) {
              return Optional.of(
                  "execution.mode.writeMode=STREAMING_WRITE supports "
                      + GridGrindContractText.streamingWriteMutationActionTypePhrase()
                      + " only; unsupported mutation action type: "
                      + mutationStep.action().actionType());
            }
            if (mutationStep.action() instanceof MutationAction.EnsureSheet) {
              seenEnsureSheet = true;
            }
            if (mutationStep.action() instanceof MutationAction.AppendRow && !seenEnsureSheet) {
              return Optional.of(
                  "execution.mode.writeMode=STREAMING_WRITE requires ENSURE_SHEET before APPEND_ROW"
                      + " so the streaming writer has a materialized sheet target");
            }
          }
          case AssertionStep _ -> {
            if (!seenEnsureSheet) {
              return Optional.of(
                  "execution.mode.writeMode=STREAMING_WRITE requires ENSURE_SHEET before any"
                      + " assertion step so the streaming workbook can materialize a sheet");
            }
          }
          case InspectionStep _ -> {
            if (!seenEnsureSheet) {
              return Optional.of(
                  "execution.mode.writeMode=STREAMING_WRITE requires ENSURE_SHEET before any"
                      + " inspection step so the streaming workbook can materialize a sheet");
            }
          }
        }
      }
      if (!seenEnsureSheet) {
        return Optional.of(
            "execution.mode.writeMode=STREAMING_WRITE requires at least one ENSURE_SHEET mutation");
      }
    }
    return Optional.empty();
  }

  static boolean directEventReadEligible(
      WorkbookPlan request, ExecutionModeSelection executionModes) {
    return executionModes.readMode() == ExecutionModeInput.ReadMode.EVENT_READ
        && executionModes.writeMode() == ExecutionModeInput.WriteMode.FULL_XSSF
        && CalculationPolicyExecutor.allowsEventRead(request.calculationPolicy())
        && request.steps().stream().allMatch(InspectionStep.class::isInstance)
        && request.persistence() instanceof WorkbookPlan.WorkbookPersistence.None
        && request.source() instanceof WorkbookPlan.WorkbookSource.ExistingFile;
  }

  static ExecutionModeSelection executionModes(WorkbookPlan request) {
    ExecutionModeInput executionMode = request.effectiveExecutionMode();
    return new ExecutionModeSelection(executionMode.readMode(), executionMode.writeMode());
  }

  private static boolean shouldExecuteCalculationBeforeStep(
      WorkbookPlan request, WorkbookStep step) {
    return CalculationPolicyExecutor.requiresMutationPrefix(request.calculationPolicy())
        && !(step instanceof MutationStep);
  }

  private CalculationExecutionOutcome executeCalculationPolicy(
      ExcelWorkbook workbook, WorkbookPlan request, ExecutionJournalRecorder journal) {
    Objects.requireNonNull(workbook, "workbook must not be null");
    Objects.requireNonNull(request, "request must not be null");
    Objects.requireNonNull(journal, "journal must not be null");
    CalculationPolicyInput policy = request.calculationPolicy();
    CalculationPolicyInput effectivePolicy = CalculationPolicyExecutor.normalize(policy);
    if (effectivePolicy.isDefault()) {
      return new CalculationExecutionOutcome(
          CalculationPolicyExecutor.notRequestedReport(effectivePolicy), null);
    }

    CalculationReport.Preflight preflightReport = null;
    int evaluationTargetCount = 0;
    if (!(effectivePolicy.effectiveStrategy() instanceof CalculationStrategyInput.DoNotCalculate)
        && !(effectivePolicy.effectiveStrategy()
            instanceof CalculationStrategyInput.ClearCachesOnly)) {
      ExecutionJournalRecorder.PhaseHandle preflightPhase = journal.beginCalculationPreflight();
      CalculationPolicyExecutor.PreflightOutcome preflight =
          CalculationPolicyExecutor.preflight(workbook, effectivePolicy);
      preflightReport = preflight.report();
      evaluationTargetCount = preflight.evaluationTargetCount();
      if (preflight.failure() != null) {
        GridGrindResponse.Problem problem = calculationProblemFor(request, preflight.failure());
        preflightPhase.fail("failed (" + problem.code() + ")");
        return new CalculationExecutionOutcome(
            CalculationPolicyExecutor.report(
                effectivePolicy,
                preflightReport,
                new CalculationReport.Execution(
                    CalculationExecutionStatus.FAILED,
                    0,
                    false,
                    false,
                    preflight.failure().message())),
            problem);
      }
      preflightPhase.succeed();
    }

    ExecutionJournalRecorder.PhaseHandle executionPhase = journal.beginCalculationExecution();
    CalculationPolicyExecutor.ExecutionOutcome execution =
        CalculationPolicyExecutor.execute(workbook, effectivePolicy, evaluationTargetCount);
    CalculationReport report =
        CalculationPolicyExecutor.report(effectivePolicy, preflightReport, execution.report());
    if (execution.failure() != null) {
      GridGrindResponse.Problem problem = calculationProblemFor(request, execution.failure());
      executionPhase.fail("failed (" + problem.code() + ")");
      return new CalculationExecutionOutcome(report, problem);
    }
    executionPhase.succeed();
    return new CalculationExecutionOutcome(report, null);
  }

  private CalculationExecutionOutcome executeStreamingCalculationPolicy(
      ExcelStreamingWorkbookWriter writer, WorkbookPlan request, ExecutionJournalRecorder journal) {
    Objects.requireNonNull(writer, "writer must not be null");
    Objects.requireNonNull(request, "request must not be null");
    Objects.requireNonNull(journal, "journal must not be null");
    CalculationPolicyInput policy = request.calculationPolicy();
    CalculationPolicyInput effectivePolicy = CalculationPolicyExecutor.normalize(policy);
    if (effectivePolicy.isDefault()) {
      return new CalculationExecutionOutcome(
          CalculationPolicyExecutor.notRequestedReport(effectivePolicy), null);
    }
    ExecutionJournalRecorder.PhaseHandle executionPhase = journal.beginCalculationExecution();
    try {
      streamingCalculationApplier.apply(writer);
      executionPhase.succeed();
      return new CalculationExecutionOutcome(
          CalculationPolicyExecutor.report(
              effectivePolicy,
              null,
              new CalculationReport.Execution(
                  CalculationExecutionStatus.SUCCEEDED, 0, false, true, null)),
          null);
    } catch (RuntimeException exception) {
      GridGrindResponse.Problem problem =
          problemFor(exception, calculationContextFor(request, "CALCULATION_EXECUTION", null));
      executionPhase.fail("failed (" + problem.code() + ")");
      return new CalculationExecutionOutcome(
          CalculationPolicyExecutor.report(
              effectivePolicy,
              null,
              new CalculationReport.Execution(
                  CalculationExecutionStatus.FAILED,
                  0,
                  false,
                  false,
                  GridGrindProblems.messageFor(exception))),
          problem);
    }
  }

  private GridGrindResponse.Problem calculationProblemFor(
      WorkbookPlan request, CalculationPolicyExecutor.FailureDetail failure) {
    if (failure.exception() != null) {
      return problemFor(
          failure.exception(), calculationContextFor(request, stageFor(failure.phase()), failure));
    }
    return GridGrindProblems.problem(
        Objects.requireNonNull(failure.code(), "failure.code must not be null"),
        failure.message(),
        calculationContextFor(request, stageFor(failure.phase()), failure),
        (Throwable) null);
  }

  private GridGrindResponse.ProblemContext.ExecuteCalculation calculationContextFor(
      WorkbookPlan request, String phase, CalculationPolicyExecutor.FailureDetail failure) {
    return new GridGrindResponse.ProblemContext.ExecuteCalculation(
        reqSourceType(request),
        reqPersistenceType(request),
        phase,
        failure == null ? null : failure.sheetName(),
        failure == null ? null : failure.address(),
        failure == null ? null : failure.formula());
  }

  private static String stageFor(CalculationPolicyExecutor.Phase phase) {
    return switch (phase) {
      case PREFLIGHT -> "CALCULATION_PREFLIGHT";
      case EXECUTION -> "CALCULATION_EXECUTION";
    };
  }

  static void deleteIfExists(Path path) {
    deleteIfExists(path, Files::deleteIfExists);
  }

  static void deleteIfExists(Path path, PathDeleteOperation deleteOperation) {
    if (path == null) {
      return;
    }
    try {
      deleteOperation.deleteIfExists(path);
    } catch (IOException ignored) {
      // Best-effort cleanup for internal temporary files only.
    }
  }

  record ExecutionModeSelection(
      ExecutionModeInput.ReadMode readMode, ExecutionModeInput.WriteMode writeMode) {}

  static ExcelOoxmlPersistenceOptions persistenceOptions(
      WorkbookPlan.WorkbookPersistence persistence) {
    return switch (persistence) {
      case WorkbookPlan.WorkbookPersistence.None _ -> new ExcelOoxmlPersistenceOptions(null, null);
      case WorkbookPlan.WorkbookPersistence.OverwriteSource overwrite ->
          OoxmlPackageSecurityConverter.toExcelPersistenceOptions(overwrite.security());
      case WorkbookPlan.WorkbookPersistence.SaveAs saveAs ->
          OoxmlPackageSecurityConverter.toExcelPersistenceOptions(saveAs.security());
    };
  }

  static ExcelOoxmlPackageSecuritySnapshot sourcePackageSecurity(
      WorkbookPlan.WorkbookSource source) {
    return switch (source) {
      case WorkbookPlan.WorkbookSource.New _ -> ExcelOoxmlPackageSecuritySnapshot.none();
      case WorkbookPlan.WorkbookSource.ExistingFile _ -> ExcelOoxmlPackageSecuritySnapshot.none();
    };
  }

  static String sourceEncryptionPassword(WorkbookPlan.WorkbookSource source) {
    return switch (source) {
      case WorkbookPlan.WorkbookSource.New _ -> null;
      case WorkbookPlan.WorkbookSource.ExistingFile existingFile ->
          existingFile.security() == null ? null : existingFile.security().password();
    };
  }

  ExcelWorkbook openWorkbook(
      WorkbookPlan.WorkbookSource source, FormulaEnvironmentInput formulaEnvironment)
      throws IOException {
    return switch (source) {
      case WorkbookPlan.WorkbookSource.New _ ->
          ExcelWorkbook.create(
              FormulaEnvironmentConverter.toExcelFormulaEnvironment(formulaEnvironment));
      case WorkbookPlan.WorkbookSource.ExistingFile existingFile ->
          ExcelWorkbook.open(
              normalizePath(existingFile.path()),
              FormulaEnvironmentConverter.toExcelFormulaEnvironment(formulaEnvironment),
              OoxmlPackageSecurityConverter.toExcelOpenOptions(existingFile.security()));
    };
  }

  GridGrindResponse.PersistenceOutcome persistWorkbook(
      ExcelWorkbook workbook,
      WorkbookPlan.WorkbookSource source,
      WorkbookPlan.WorkbookPersistence persistence)
      throws IOException {
    Objects.requireNonNull(workbook, "workbook must not be null");
    return switch (persistence) {
      case WorkbookPlan.WorkbookPersistence.None _ ->
          new GridGrindResponse.PersistenceOutcome.NotSaved();
      case WorkbookPlan.WorkbookPersistence.SaveAs saveAs -> {
        Path executionPath = normalizePath(saveAs.path());
        workbook.save(executionPath, persistenceOptions(saveAs), tempFileFactory::createTempFile);
        yield new GridGrindResponse.PersistenceOutcome.SavedAs(
            saveAs.path(), executionPath.toString());
      }
      case WorkbookPlan.WorkbookPersistence.OverwriteSource overwrite -> {
        if (!(source instanceof WorkbookPlan.WorkbookSource.ExistingFile existingFile)) {
          throw new IllegalArgumentException("OVERWRITE persistence requires an EXISTING source");
        }
        Path executionPath = normalizePath(existingFile.path());
        workbook.save(
            executionPath, persistenceOptions(overwrite), tempFileFactory::createTempFile);
        yield new GridGrindResponse.PersistenceOutcome.Overwritten(
            existingFile.path(), executionPath.toString());
      }
    };
  }

  GridGrindResponse.PersistenceOutcome persistStreamingWorkbook(
      Path materializedPath,
      WorkbookPlan.WorkbookPersistence persistence,
      WorkbookPlan.WorkbookSource source)
      throws IOException {
    Objects.requireNonNull(materializedPath, "materializedPath must not be null");
    return switch (persistence) {
      case WorkbookPlan.WorkbookPersistence.None _ ->
          new GridGrindResponse.PersistenceOutcome.NotSaved();
      case WorkbookPlan.WorkbookPersistence.SaveAs saveAs -> {
        Path executionPath = normalizePath(saveAs.path());
        ExcelOoxmlPackageSecuritySupport.persistMaterializedWorkbook(
            materializedPath,
            executionPath,
            sourcePackageSecurity(source),
            sourceEncryptionPassword(source),
            true,
            persistenceOptions(saveAs));
        yield new GridGrindResponse.PersistenceOutcome.SavedAs(
            saveAs.path(), executionPath.toString());
      }
      case WorkbookPlan.WorkbookPersistence.OverwriteSource overwrite -> {
        if (!(source instanceof WorkbookPlan.WorkbookSource.ExistingFile existingFile)) {
          throw new IllegalArgumentException("OVERWRITE persistence requires an EXISTING source");
        }
        Path executionPath = normalizePath(existingFile.path());
        ExcelOoxmlPackageSecuritySupport.persistMaterializedWorkbook(
            materializedPath,
            executionPath,
            sourcePackageSecurity(source),
            sourceEncryptionPassword(source),
            true,
            persistenceOptions(overwrite));
        yield new GridGrindResponse.PersistenceOutcome.Overwritten(
            existingFile.path(), executionPath.toString());
      }
    };
  }

  static String persistencePath(
      WorkbookPlan.WorkbookSource source, WorkbookPlan.WorkbookPersistence persistence) {
    return normalizedPersistencePath(source, persistence);
  }

  static WorkbookLocation workbookLocationFor(
      WorkbookPlan.WorkbookSource source, WorkbookPlan.WorkbookPersistence persistence) {
    String persistencePath = normalizedPersistencePath(source, persistence);
    if (persistencePath != null) {
      return new WorkbookLocation.StoredWorkbook(Path.of(persistencePath));
    }
    return switch (source) {
      case WorkbookPlan.WorkbookSource.New _ -> new WorkbookLocation.UnsavedWorkbook();
      case WorkbookPlan.WorkbookSource.ExistingFile existingFile ->
          new WorkbookLocation.StoredWorkbook(normalizePath(existingFile.path()));
    };
  }

  private GridGrindResponse closeWorkbook(
      ExcelWorkbook workbook,
      GridGrindResponse response,
      WorkbookPlan request,
      ExecutionJournalRecorder journal,
      GridGrindProblemCode primaryFailureCode,
      Integer failedStepIndex,
      String failedStepId) {
    ExecutionJournalRecorder.PhaseHandle closePhase = journal.beginClose();
    try {
      workbookCloser.close(workbook);
      closePhase.succeed();
      return switch (response) {
        case GridGrindResponse.Success success ->
            new GridGrindResponse.Success(
                success.protocolVersion(),
                journal.buildSuccess(request.steps().size()),
                success.calculation(),
                success.persistence(),
                success.warnings(),
                success.assertions(),
                success.inspections());
        case GridGrindResponse.Failure failure ->
            new GridGrindResponse.Failure(
                failure.protocolVersion(),
                journal.buildFailure(
                    request.steps().size(), primaryFailureCode, failedStepIndex, failedStepId),
                failure.calculation(),
                failure.problem());
      };
    } catch (Exception closeFailure) {
      GridGrindResponse.Problem closeProblem =
          problemFor(
              closeFailure,
              new GridGrindResponse.ProblemContext.ExecuteRequest(
                  reqSourceType(request), reqPersistenceType(request)));
      closePhase.fail("failed (" + closeProblem.code() + ")");
      if (response instanceof GridGrindResponse.Failure existingFailure) {
        return new GridGrindResponse.Failure(
            existingFailure.protocolVersion(),
            journal.buildFailure(
                request.steps().size(), primaryFailureCode, failedStepIndex, failedStepId),
            existingFailure.calculation(),
            GridGrindProblems.appendCause(
                existingFailure.problem(), GridGrindProblems.problemCause(closeProblem)));
      }
      return new GridGrindResponse.Failure(
          request.protocolVersion(),
          journal.buildFailure(request.steps().size(), closeProblem.code(), null, null),
          response.calculation(),
          closeProblem);
    }
  }

  private GridGrindResponse closeReadableWorkbook(
      ExcelOoxmlPackageSecuritySupport.ReadableWorkbook workbook,
      GridGrindResponse response,
      WorkbookPlan request,
      ExecutionJournalRecorder journal,
      GridGrindProblemCode primaryFailureCode,
      Integer failedStepIndex,
      String failedStepId) {
    if (workbook == null) {
      return response;
    }
    ExecutionJournalRecorder.PhaseHandle closePhase = journal.beginClose();
    try {
      readableWorkbookCloser.close(workbook);
      closePhase.succeed();
      return switch (response) {
        case GridGrindResponse.Success success ->
            new GridGrindResponse.Success(
                success.protocolVersion(),
                journal.buildSuccess(request.steps().size()),
                success.calculation(),
                success.persistence(),
                success.warnings(),
                success.assertions(),
                success.inspections());
        case GridGrindResponse.Failure failure ->
            new GridGrindResponse.Failure(
                failure.protocolVersion(),
                journal.buildFailure(
                    request.steps().size(), primaryFailureCode, failedStepIndex, failedStepId),
                failure.calculation(),
                failure.problem());
      };
    } catch (Exception closeFailure) {
      GridGrindResponse.Problem closeProblem =
          problemFor(
              closeFailure,
              new GridGrindResponse.ProblemContext.ExecuteRequest(
                  reqSourceType(request), reqPersistenceType(request)));
      closePhase.fail("failed (" + closeProblem.code() + ")");
      if (response instanceof GridGrindResponse.Failure existingFailure) {
        return new GridGrindResponse.Failure(
            existingFailure.protocolVersion(),
            journal.buildFailure(
                request.steps().size(), primaryFailureCode, failedStepIndex, failedStepId),
            existingFailure.calculation(),
            GridGrindProblems.appendCause(
                existingFailure.problem(), GridGrindProblems.problemCause(closeProblem)));
      }
      return new GridGrindResponse.Failure(
          request.protocolVersion(),
          journal.buildFailure(request.steps().size(), closeProblem.code(), null, null),
          response.calculation(),
          closeProblem);
    }
  }

  GridGrindResponse guardUnexpectedRuntime(
      GridGrindProtocolVersion protocolVersion,
      WorkbookPlan request,
      ExecutionJournalRecorder journal,
      java.util.function.Supplier<GridGrindResponse> workflow) {
    try {
      return workflow.get();
    } catch (RuntimeException exception) {
      GridGrindResponse.Problem problem =
          problemFor(
              exception,
              new GridGrindResponse.ProblemContext.ExecuteRequest(
                  reqSourceType(request), reqPersistenceType(request)));
      return failureResponse(
          protocolVersion,
          journal,
          request.steps().size(),
          CalculationPolicyExecutor.notRequestedReport(request.calculationPolicy()),
          problem,
          null,
          null);
    }
  }

  GridGrindResponse guardUnexpectedRuntime(
      GridGrindProtocolVersion protocolVersion,
      WorkbookPlan request,
      ExecutionJournalRecorder journal,
      ExcelWorkbook workbook,
      java.util.function.Supplier<GridGrindResponse> workflow) {
    try {
      return workflow.get();
    } catch (RuntimeException exception) {
      GridGrindResponse.Problem problem =
          problemFor(
              exception,
              new GridGrindResponse.ProblemContext.ExecuteRequest(
                  reqSourceType(request), reqPersistenceType(request)));
      return closeWorkbook(
          workbook,
          failureResponse(
              protocolVersion,
              journal,
              request.steps().size(),
              CalculationPolicyExecutor.notRequestedReport(request.calculationPolicy()),
              problem,
              null,
              null),
          request,
          journal,
          problem.code(),
          null,
          null);
    }
  }

  static GridGrindProblemCode problemCodeFor(Throwable exception) {
    return GridGrindProblems.codeFor(exception);
  }

  static GridGrindResponse.ProblemContext enrichContext(
      GridGrindResponse.ProblemContext context, Throwable exception) {
    return GridGrindProblems.enrichContext(context, exception);
  }

  private static GridGrindResponse.Problem problemFor(
      Throwable exception, GridGrindResponse.ProblemContext context) {
    return GridGrindProblems.fromException(exception, context);
  }

  static String reqSourceType(WorkbookPlan request) {
    if (request.source() instanceof WorkbookPlan.WorkbookSource.New) {
      return "NEW";
    }
    return "EXISTING";
  }

  static String reqPersistenceType(WorkbookPlan request) {
    return switch (request.persistence()) {
      case WorkbookPlan.WorkbookPersistence.None _ -> "NONE";
      case WorkbookPlan.WorkbookPersistence.OverwriteSource _ -> "OVERWRITE";
      case WorkbookPlan.WorkbookPersistence.SaveAs _ -> "SAVE_AS";
    };
  }

  static String reqSourcePath(WorkbookPlan request) {
    return switch (request.source()) {
      case WorkbookPlan.WorkbookSource.New _ -> null;
      case WorkbookPlan.WorkbookSource.ExistingFile existingFile ->
          normalizePath(existingFile.path()).toString();
    };
  }

  private static String normalizedPersistencePath(
      WorkbookPlan.WorkbookSource source, WorkbookPlan.WorkbookPersistence persistence) {
    return switch (persistence) {
      case WorkbookPlan.WorkbookPersistence.None _ -> null;
      case WorkbookPlan.WorkbookPersistence.SaveAs saveAs ->
          normalizePath(saveAs.path()).toString();
      case WorkbookPlan.WorkbookPersistence.OverwriteSource _ ->
          source instanceof WorkbookPlan.WorkbookSource.ExistingFile existingFile
              ? normalizePath(existingFile.path()).toString()
              : null;
    };
  }

  private static Path normalizePath(String path) {
    return Path.of(path).toAbsolutePath().normalize();
  }

  private static GridGrindResponse.Failure failureResponse(
      GridGrindProtocolVersion protocolVersion,
      ExecutionJournalRecorder journal,
      int plannedStepCount,
      GridGrindResponse.Problem problem,
      Integer failedStepIndex,
      String failedStepId) {
    return failureResponse(
        protocolVersion, journal, plannedStepCount, null, problem, failedStepIndex, failedStepId);
  }

  private static GridGrindResponse.Failure failureResponse(
      GridGrindProtocolVersion protocolVersion,
      ExecutionJournalRecorder journal,
      int plannedStepCount,
      CalculationReport calculation,
      GridGrindResponse.Problem problem,
      Integer failedStepIndex,
      String failedStepId) {
    return new GridGrindResponse.Failure(
        protocolVersion,
        journal.buildFailure(plannedStepCount, problem.code(), failedStepIndex, failedStepId),
        calculation,
        problem);
  }

  static String stepType(WorkbookStep step) {
    return switch (step) {
      case MutationStep mutationStep -> mutationStep.action().actionType();
      case AssertionStep assertionStep -> assertionStep.assertion().assertionType();
      case InspectionStep inspectionStep -> inspectionStep.query().queryType();
    };
  }

  static String sheetNameFor(WorkbookStep step) {
    return switch (step) {
      case MutationStep mutationStep -> {
        String fromAction = sheetNameFor(mutationStep.action());
        yield fromAction != null ? fromAction : sheetNameFor(mutationStep.target());
      }
      case AssertionStep assertionStep -> sheetNameFor(assertionStep.target());
      case InspectionStep inspectionStep -> sheetNameFor(inspectionStep.target());
    };
  }

  static String sheetNameFor(WorkbookStep step, Exception exception) {
    String fromStep = sheetNameFor(step);
    return fromStep != null ? fromStep : sheetNameFor(exception);
  }

  static String addressFor(WorkbookStep step) {
    return switch (step) {
      case MutationStep mutationStep -> {
        String fromAction = addressFor(mutationStep.action());
        yield fromAction != null ? fromAction : addressFor(mutationStep.target());
      }
      case AssertionStep assertionStep -> addressFor(assertionStep.target());
      case InspectionStep inspectionStep -> addressFor(inspectionStep.target());
    };
  }

  static String addressFor(WorkbookStep step, Exception exception) {
    String fromStep = addressFor(step);
    return fromStep != null ? fromStep : addressFor(exception);
  }

  static String rangeFor(WorkbookStep step) {
    return switch (step) {
      case MutationStep mutationStep -> {
        String fromAction = rangeFor(mutationStep.action());
        yield fromAction != null ? fromAction : rangeFor(mutationStep.target());
      }
      case AssertionStep assertionStep -> rangeFor(assertionStep.target());
      case InspectionStep inspectionStep -> rangeFor(inspectionStep.target());
    };
  }

  static String rangeFor(WorkbookStep step, Exception exception) {
    String fromStep = rangeFor(step);
    return fromStep != null ? fromStep : rangeFor(exception);
  }

  static String formulaFor(WorkbookStep step) {
    return switch (step) {
      case MutationStep mutationStep -> formulaFor(mutationStep.action());
      case AssertionStep assertionStep -> formulaFor(assertionStep.assertion());
      case InspectionStep _ -> null;
    };
  }

  static String formulaFor(WorkbookStep step, Exception exception) {
    String fromStep = formulaFor(step);
    return fromStep != null ? fromStep : formulaFor(exception);
  }

  static String namedRangeNameFor(WorkbookStep step) {
    return switch (step) {
      case MutationStep mutationStep -> {
        String fromAction = namedRangeNameFor(mutationStep.action());
        yield fromAction != null ? fromAction : namedRangeNameFor(mutationStep.target());
      }
      case AssertionStep assertionStep -> namedRangeNameFor(assertionStep.target());
      case InspectionStep inspectionStep -> namedRangeNameFor(inspectionStep.target());
    };
  }

  static String namedRangeNameFor(WorkbookStep step, Exception exception) {
    String fromStep = namedRangeNameFor(step);
    return fromStep != null ? fromStep : namedRangeNameFor(exception);
  }

  static String sheetNameFor(Exception exception) {
    return switch (exception) {
      case FormulaException formulaException -> formulaException.sheetName();
      case SheetNotFoundException sheetNotFoundException -> sheetNotFoundException.sheetName();
      case NamedRangeNotFoundException namedRangeNotFoundException ->
          switch (namedRangeNotFoundException.scope()) {
            case ExcelNamedRangeScope.WorkbookScope _ -> null;
            case ExcelNamedRangeScope.SheetScope sheetScope -> sheetScope.sheetName();
          };
      default -> null;
    };
  }

  static String addressFor(Exception exception) {
    return switch (exception) {
      case FormulaException formulaException -> formulaException.address();
      case CellNotFoundException cellNotFoundException -> cellNotFoundException.address();
      case InvalidCellAddressException invalidCellAddressException ->
          invalidCellAddressException.address();
      default -> null;
    };
  }

  static String rangeFor(Exception exception) {
    if (exception instanceof InvalidRangeAddressException invalidRangeAddressException) {
      return invalidRangeAddressException.range();
    }
    return null;
  }

  static String formulaFor(Exception exception) {
    if (exception instanceof FormulaException formulaException) {
      return formulaException.formula();
    }
    return null;
  }

  static String namedRangeNameFor(Exception exception) {
    if (exception instanceof NamedRangeNotFoundException namedRangeNotFoundException) {
      return namedRangeNotFoundException.name();
    }
    return null;
  }

  static String sheetNameFor(MutationAction action) {
    return switch (action) {
      case MutationAction.SetTable setTable -> setTable.table().sheetName();
      case MutationAction.SetPivotTable setPivotTable -> setPivotTable.pivotTable().sheetName();
      case MutationAction.SetNamedRange setNamedRange ->
          setNamedRange.target().sheetName() != null
              ? setNamedRange.target().sheetName()
              : switch (setNamedRange.scope()) {
                case NamedRangeScope.Workbook _ -> null;
                case NamedRangeScope.Sheet sheet -> sheet.sheetName();
              };
      default -> null;
    };
  }

  static String addressFor(MutationAction action) {
    if (action instanceof MutationAction.SetPivotTable setPivotTable) {
      return setPivotTable.pivotTable().anchor().topLeftAddress();
    }
    return null;
  }

  static String rangeFor(MutationAction action) {
    return switch (action) {
      case MutationAction.SetTable setTable -> setTable.table().range();
      case MutationAction.SetPivotTable setPivotTable ->
          switch (setPivotTable.pivotTable().source()) {
            case PivotTableInput.Source.Range range -> range.range();
            case PivotTableInput.Source.NamedRange _ -> null;
            case PivotTableInput.Source.Table _ -> null;
          };
      case MutationAction.SetNamedRange setNamedRange -> setNamedRange.target().range();
      case MutationAction.SetConditionalFormatting setConditionalFormatting ->
          setConditionalFormatting.conditionalFormatting().ranges().size() == 1
              ? setConditionalFormatting.conditionalFormatting().ranges().getFirst()
              : null;
      default -> null;
    };
  }

  static String formulaFor(MutationAction action) {
    return switch (action) {
      case MutationAction.SetCell setCell ->
          setCell.value() instanceof CellInput.Formula formula ? inlineFormula(formula) : null;
      case MutationAction.SetNamedRange setNamedRange -> setNamedRange.target().formula();
      default -> null;
    };
  }

  private static String inlineFormula(CellInput.Formula formula) {
    if (formula.source() instanceof TextSourceInput.Inline inline) {
      return inline.text();
    }
    return null;
  }

  static String formulaFor(Assertion assertion) {
    if (assertion instanceof Assertion.FormulaText formulaText) {
      return formulaText.formula();
    }
    return null;
  }

  static String namedRangeNameFor(MutationAction action) {
    return switch (action) {
      case MutationAction.SetNamedRange setNamedRange -> setNamedRange.name();
      case MutationAction.SetPivotTable setPivotTable ->
          switch (setPivotTable.pivotTable().source()) {
            case PivotTableInput.Source.NamedRange namedRange -> namedRange.name();
            case PivotTableInput.Source.Range _ -> null;
            case PivotTableInput.Source.Table _ -> null;
          };
      default -> null;
    };
  }

  static String sheetNameFor(Selector selector) {
    return switch (selector) {
      case WorkbookSelector _ -> null;
      case SheetSelector.All _ -> null;
      case SheetSelector.ByName byName -> byName.name();
      case SheetSelector.ByNames byNames ->
          byNames.names().size() == 1 ? byNames.names().getFirst() : null;
      case CellSelector cellSelector -> singleSheetName(cellSelector);
      case RangeSelector rangeSelector -> singleSheetName(rangeSelector);
      case RowBandSelector.Span span -> span.sheetName();
      case RowBandSelector.Insertion insertion -> insertion.sheetName();
      case ColumnBandSelector.Span span -> span.sheetName();
      case ColumnBandSelector.Insertion insertion -> insertion.sheetName();
      case DrawingObjectSelector.AllOnSheet allOnSheet -> allOnSheet.sheetName();
      case DrawingObjectSelector.ByName byName -> byName.sheetName();
      case ChartSelector.AllOnSheet allOnSheet -> allOnSheet.sheetName();
      case ChartSelector.ByName byName -> byName.sheetName();
      case TableSelector.All _ -> null;
      case TableSelector.ByName _ -> null;
      case TableSelector.ByNames _ -> null;
      case TableSelector.ByNameOnSheet byNameOnSheet -> byNameOnSheet.sheetName();
      case PivotTableSelector.All _ -> null;
      case PivotTableSelector.ByName _ -> null;
      case PivotTableSelector.ByNames _ -> null;
      case PivotTableSelector.ByNameOnSheet byNameOnSheet -> byNameOnSheet.sheetName();
      case NamedRangeSelector namedRangeSelector -> singleSheetName(namedRangeSelector);
      case TableRowSelector.AllRows allRows -> sheetNameFor(allRows.table());
      case TableRowSelector.ByIndex byIndex -> sheetNameFor(byIndex.table());
      case TableRowSelector.ByKeyCell byKeyCell -> sheetNameFor(byKeyCell.table());
      case TableCellSelector.ByColumnName byColumnName -> sheetNameFor(byColumnName.row());
      default -> null;
    };
  }

  static String addressFor(Selector selector) {
    return switch (selector) {
      case CellSelector.ByAddress byAddress -> byAddress.address();
      case CellSelector.ByQualifiedAddresses qualifiedAddresses ->
          qualifiedAddresses.cells().size() == 1
              ? qualifiedAddresses.cells().getFirst().address()
              : null;
      case RangeSelector.RectangularWindow window -> window.topLeftAddress();
      default -> null;
    };
  }

  static String rangeFor(Selector selector) {
    return switch (selector) {
      case RangeSelector.ByRange byRange -> byRange.range();
      case RangeSelector.ByRanges byRanges ->
          byRanges.ranges().size() == 1 ? byRanges.ranges().getFirst() : null;
      case RangeSelector.RectangularWindow window -> window.range();
      default -> null;
    };
  }

  static String namedRangeNameFor(Selector selector) {
    if (selector instanceof NamedRangeSelector namedRangeSelector) {
      return singleNamedRangeName(namedRangeSelector);
    }
    return null;
  }

  static String singleSheetName(CellSelector selector) {
    return switch (selector) {
      case CellSelector.AllUsedInSheet allUsedInSheet -> allUsedInSheet.sheetName();
      case CellSelector.ByAddress byAddress -> byAddress.sheetName();
      case CellSelector.ByAddresses byAddresses -> byAddresses.sheetName();
      case CellSelector.ByQualifiedAddresses qualifiedAddresses ->
          qualifiedAddresses.cells().stream()
                      .map(CellSelector.QualifiedAddress::sheetName)
                      .distinct()
                      .count()
                  == 1
              ? qualifiedAddresses.cells().getFirst().sheetName()
              : null;
    };
  }

  static String singleSheetName(RangeSelector selector) {
    return switch (selector) {
      case RangeSelector.AllOnSheet allOnSheet -> allOnSheet.sheetName();
      case RangeSelector.ByRange byRange -> byRange.sheetName();
      case RangeSelector.ByRanges byRanges -> byRanges.sheetName();
      case RangeSelector.RectangularWindow window -> window.sheetName();
    };
  }

  static String singleSheetName(NamedRangeSelector selector) {
    return switch (selector) {
      case NamedRangeSelector.All _ -> null;
      case NamedRangeSelector.ByName _ -> null;
      case NamedRangeSelector.ByNames _ -> null;
      case NamedRangeSelector.WorkbookScope _ -> null;
      case NamedRangeSelector.SheetScope sheetScope -> sheetScope.sheetName();
      case NamedRangeSelector.AnyOf anyOf ->
          anyOf.selectors().size() == 1 ? singleSheetName(anyOf.selectors().getFirst()) : null;
    };
  }

  static String singleSheetName(NamedRangeSelector.Ref selector) {
    return switch (selector) {
      case NamedRangeSelector.ByName _ -> null;
      case NamedRangeSelector.WorkbookScope _ -> null;
      case NamedRangeSelector.SheetScope sheetScope -> sheetScope.sheetName();
    };
  }

  static String singleNamedRangeName(NamedRangeSelector selector) {
    return switch (selector) {
      case NamedRangeSelector.All _ -> null;
      case NamedRangeSelector.ByName byName -> byName.name();
      case NamedRangeSelector.ByNames byNames ->
          byNames.names().size() == 1 ? byNames.names().getFirst() : null;
      case NamedRangeSelector.WorkbookScope workbookScope -> workbookScope.name();
      case NamedRangeSelector.SheetScope sheetScope -> sheetScope.name();
      case NamedRangeSelector.AnyOf anyOf ->
          anyOf.selectors().size() == 1 ? singleNamedRangeName(anyOf.selectors().getFirst()) : null;
    };
  }

  static String singleNamedRangeName(NamedRangeSelector.Ref selector) {
    return switch (selector) {
      case NamedRangeSelector.ByName byName -> byName.name();
      case NamedRangeSelector.WorkbookScope workbookScope -> workbookScope.name();
      case NamedRangeSelector.SheetScope sheetScope -> sheetScope.name();
    };
  }

  /** Functional interface for creating executor-owned temporary workbook files. */
  @FunctionalInterface
  interface TempFileFactory {
    /** Creates one temporary workbook file path for internal execution workflows. */
    Path createTempFile(String prefix, String suffix) throws IOException;
  }

  /** Closes one materialized event-read workbook path owned by this executor. */
  @FunctionalInterface
  interface ReadableWorkbookCloser {
    /** Closes one event-read materialized workbook owned by the executor. */
    void close(ExcelOoxmlPackageSecuritySupport.ReadableWorkbook workbook) throws IOException;
  }

  /** Best-effort filesystem deletion seam used by executor-owned cleanup tests. */
  @FunctionalInterface
  interface PathDeleteOperation {
    /** Deletes one executor-owned path when present. */
    void deleteIfExists(Path path) throws IOException;
  }

  /** Functional interface for closing an ExcelWorkbook after request execution. */
  @FunctionalInterface
  interface WorkbookCloser {
    /** Closes the workbook, releasing any held resources. */
    void close(ExcelWorkbook workbook) throws IOException;
  }

  /** Applies the streaming-mode calculation side effect at the execution boundary. */
  @FunctionalInterface
  interface StreamingCalculationApplier {
    /** Applies streaming-mode calculation state to the low-memory workbook writer. */
    void apply(ExcelStreamingWorkbookWriter writer);
  }

  private record CalculationExecutionOutcome(
      CalculationReport report, GridGrindResponse.Problem failure) {
    private CalculationExecutionOutcome {
      Objects.requireNonNull(report, "report must not be null");
    }
  }
}
