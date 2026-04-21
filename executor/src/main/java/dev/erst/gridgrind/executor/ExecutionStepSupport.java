package dev.erst.gridgrind.executor;

import dev.erst.gridgrind.contract.assertion.AssertionResult;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.contract.step.MutationStep;
import dev.erst.gridgrind.contract.step.WorkbookStep;
import dev.erst.gridgrind.excel.ExcelEventWorkbookReader;
import dev.erst.gridgrind.excel.ExcelStreamingWorkbookWriter;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookCommand;
import dev.erst.gridgrind.excel.WorkbookCommandExecutor;
import dev.erst.gridgrind.excel.WorkbookLocation;
import dev.erst.gridgrind.excel.WorkbookReadCommand;
import dev.erst.gridgrind.excel.WorkbookReadExecutor;
import java.io.IOException;
import java.nio.file.Path;
import java.util.List;
import java.util.Objects;

/** Executes workbook steps and materialized-read seams for the request executor. */
final class ExecutionStepSupport {
  private final WorkbookCommandExecutor commandExecutor;
  private final WorkbookReadExecutor readExecutor;
  private final SemanticSelectorResolver selectorResolver;
  private final AssertionExecutor assertionExecutor;
  private final TempFileFactory tempFileFactory;

  ExecutionStepSupport(
      WorkbookCommandExecutor commandExecutor,
      WorkbookReadExecutor readExecutor,
      SemanticSelectorResolver selectorResolver,
      AssertionExecutor assertionExecutor,
      TempFileFactory tempFileFactory) {
    this.commandExecutor =
        Objects.requireNonNull(commandExecutor, "commandExecutor must not be null");
    this.readExecutor = Objects.requireNonNull(readExecutor, "readExecutor must not be null");
    this.selectorResolver =
        Objects.requireNonNull(selectorResolver, "selectorResolver must not be null");
    this.assertionExecutor =
        Objects.requireNonNull(assertionExecutor, "assertionExecutor must not be null");
    this.tempFileFactory =
        Objects.requireNonNull(tempFileFactory, "tempFileFactory must not be null");
  }

  void executeMutationStep(ExcelWorkbook workbook, MutationStep mutationStep) throws IOException {
    Selector resolvedTarget =
        selectorResolver.resolveMutationTarget(
            workbook, mutationStep.target(), mutationStep.action());
    WorkbookCommand command =
        WorkbookCommandConverter.toCommand(resolvedTarget, mutationStep.action());
    commandExecutor.apply(workbook, command);
  }

  void executeStreamingMutationStep(ExcelStreamingWorkbookWriter writer, MutationStep mutationStep)
      throws IOException {
    WorkbookCommand command = WorkbookCommandConverter.toCommand(mutationStep);
    writer.apply(command);
  }

  GridGrindResponse.ProblemContext.ExecuteStep executeStepContext(
      WorkbookPlan request, int stepIndex, WorkbookStep step, Exception exception) {
    return new GridGrindResponse.ProblemContext.ExecuteStep(
        ExecutionRequestPaths.reqSourceType(request),
        ExecutionRequestPaths.reqPersistenceType(request),
        stepIndex,
        step.stepId(),
        step.stepKind(),
        ExecutionStepKinds.stepType(step),
        ExecutionDiagnosticFields.sheetNameFor(step, exception),
        ExecutionDiagnosticFields.addressFor(step, exception),
        ExecutionDiagnosticFields.rangeFor(step, exception),
        ExecutionDiagnosticFields.formulaFor(step, exception),
        ExecutionDiagnosticFields.namedRangeNameFor(step, exception));
  }

  GridGrindResponse.ProblemContext.ResolveInputs resolveInputsContext(
      WorkbookPlan request, Exception exception) {
    return new GridGrindResponse.ProblemContext.ResolveInputs(
        ExecutionRequestPaths.reqSourceType(request),
        ExecutionRequestPaths.reqPersistenceType(request),
        exception instanceof InputSourceException inputSourceException
            ? inputSourceException.inputKind()
            : null,
        exception instanceof InputSourceException inputSourceException
            ? inputSourceException.inputPath()
            : null);
  }

  InspectionResult executeInspectionStep(
      InspectionStep inspectionStep,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation,
      dev.erst.gridgrind.contract.dto.ExecutionModeInput.ReadMode readMode)
      throws IOException {
    return switch (readMode) {
      case FULL_XSSF -> executeFullInspectionStep(inspectionStep, workbook, workbookLocation);
      case EVENT_READ -> executeEventInspectionAgainstWorkbook(inspectionStep, workbook);
    };
  }

  AssertionResult executeAssertionStep(
      AssertionStep assertionStep,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation,
      dev.erst.gridgrind.contract.dto.ExecutionModeInput.ReadMode readMode)
      throws IOException, AssertionFailedException {
    return switch (readMode) {
      case FULL_XSSF -> assertionExecutor.execute(assertionStep, workbook, workbookLocation);
      case EVENT_READ ->
          throw new IllegalStateException(
              "executionMode.readMode=EVENT_READ does not support assertion steps");
    };
  }

  InspectionResult executeInspectionAgainstMaterializedPath(
      InspectionStep inspectionStep,
      WorkbookLocation workbookLocation,
      dev.erst.gridgrind.contract.dto.ExecutionModeInput.ReadMode readMode,
      Path materializedPath)
      throws IOException {
    return switch (readMode) {
      case FULL_XSSF ->
          executeFullInspectionAgainstMaterializedPath(
              inspectionStep, workbookLocation, materializedPath);
      case EVENT_READ -> executeEventInspection(materializedPath, inspectionStep);
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
      ExecutionWorkbookSupport.deleteIfExists(tempPath);
    }
  }

  InspectionResult executeStreamingInspectionStep(
      ExcelStreamingWorkbookWriter writer,
      InspectionStep inspectionStep,
      WorkbookLocation workbookLocation,
      dev.erst.gridgrind.contract.dto.ExecutionModeInput.ReadMode readMode)
      throws IOException {
    Path tempPath = tempFileFactory.createTempFile("gridgrind-streaming-step-", ".xlsx");
    try {
      writer.save(tempPath);
      return executeInspectionAgainstMaterializedPath(
          inspectionStep, workbookLocation, readMode, tempPath);
    } finally {
      ExecutionWorkbookSupport.deleteIfExists(tempPath);
    }
  }

  AssertionResult executeStreamingAssertionStep(
      ExcelStreamingWorkbookWriter writer,
      AssertionStep assertionStep,
      WorkbookLocation workbookLocation)
      throws IOException, AssertionFailedException {
    Path tempPath = tempFileFactory.createTempFile("gridgrind-streaming-step-", ".xlsx");
    try {
      writer.save(tempPath);
      return executeFullAssertionAgainstMaterializedPath(assertionStep, workbookLocation, tempPath);
    } finally {
      ExecutionWorkbookSupport.deleteIfExists(tempPath);
    }
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

  InspectionResult executeEventInspection(Path workbookPath, InspectionStep inspectionStep)
      throws IOException {
    ExcelEventWorkbookReader eventWorkbookReader = new ExcelEventWorkbookReader();
    WorkbookReadCommand command = InspectionCommandConverter.toReadCommand(inspectionStep);
    dev.erst.gridgrind.excel.WorkbookReadResult result =
        eventWorkbookReader
            .apply(workbookPath, List.of((WorkbookReadCommand.Introspection) command))
            .getFirst();
    return InspectionResultConverter.toReadResult(result);
  }
}
