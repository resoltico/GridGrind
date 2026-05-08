package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.assertion.AssertionResult;
import dev.erst.gridgrind.contract.catalog.GridGrindExecutionModeMetadata;
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
import dev.erst.gridgrind.excel.WorkbookExecutionEngine;
import dev.erst.gridgrind.excel.WorkbookLocation;
import dev.erst.gridgrind.excel.WorkbookReadCommand;
import java.io.IOException;
import java.nio.file.Path;
import java.util.List;
import java.util.Objects;

/** Executes workbook steps and materialized-read seams for the request executor. */
final class ExecutionStepSupport {
  private final WorkbookExecutionEngine workbookEngine;
  private final SemanticSelectorResolver selectorResolver;
  private final AssertionExecutor assertionExecutor;
  private final TempFileFactory tempFileFactory;

  ExecutionStepSupport(
      WorkbookExecutionEngine workbookEngine,
      SemanticSelectorResolver selectorResolver,
      AssertionExecutor assertionExecutor,
      TempFileFactory tempFileFactory) {
    this.workbookEngine = Objects.requireNonNull(workbookEngine, "workbookEngine must not be null");
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
    workbookEngine.apply(workbook, command);
  }

  void executeStreamingMutationStep(ExcelStreamingWorkbookWriter writer, MutationStep mutationStep)
      throws IOException {
    WorkbookCommand command = WorkbookCommandConverter.toCommand(mutationStep);
    writer.apply(command);
  }

  dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteStep executeStepContext(
      WorkbookPlan request, int stepIndex, WorkbookStep step, Exception exception) {
    return new dev.erst.gridgrind.contract.dto.ProblemContext.ExecuteStep(
        ExecutionRequestPaths.requestShape(request),
        new dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.StepReference(
            stepIndex, step.stepId(), step.stepKind(), ExecutionStepKinds.stepType(step)),
        ExecutionDiagnosticFields.locationFor(step, exception));
  }

  dev.erst.gridgrind.contract.dto.ProblemContext.ResolveInputs resolveInputsContext(
      WorkbookPlan request, Exception exception) {
    return new dev.erst.gridgrind.contract.dto.ProblemContext.ResolveInputs(
        ExecutionRequestPaths.requestShape(request),
        exception instanceof InputSourceException inputSourceException
            ? inputSourceException.inputPath() != null
                ? dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.InputReference
                    .path(inputSourceException.inputKind(), inputSourceException.inputPath())
                : dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.InputReference
                    .kind(inputSourceException.inputKind())
            : dev.erst.gridgrind.contract.dto.ProblemContextWorkbookSurfaces.InputReference
                .unknown());
  }

  InspectionResult executeInspectionStep(
      InspectionStep inspectionStep,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation,
      dev.erst.gridgrind.contract.dto.ExecutionModeInput.ReadMode readMode)
      throws IOException {
    if (readMode == dev.erst.gridgrind.contract.dto.ExecutionModeInput.ReadMode.FULL_XSSF) {
      return executeFullInspectionStep(inspectionStep, workbook, workbookLocation);
    }
    return executeEventInspectionAgainstWorkbook(inspectionStep, workbook);
  }

  AssertionResult executeAssertionStep(
      AssertionStep assertionStep,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation,
      dev.erst.gridgrind.contract.dto.ExecutionModeInput.ReadMode readMode)
      throws IOException, AssertionFailedException {
    if (readMode == dev.erst.gridgrind.contract.dto.ExecutionModeInput.ReadMode.FULL_XSSF) {
      return assertionExecutor.execute(assertionStep, workbook, workbookLocation);
    }
    throw new IllegalStateException(
        "executionMode.readMode=EVENT_READ does not support assertion steps");
  }

  InspectionResult executeInspectionAgainstMaterializedPath(
      InspectionStep inspectionStep,
      WorkbookLocation workbookLocation,
      dev.erst.gridgrind.contract.dto.ExecutionModeInput.ReadMode readMode,
      Path materializedPath)
      throws IOException {
    if (readMode == dev.erst.gridgrind.contract.dto.ExecutionModeInput.ReadMode.FULL_XSSF) {
      return executeFullInspectionAgainstMaterializedPath(
          inspectionStep, workbookLocation, materializedPath);
    }
    return executeEventInspection(materializedPath, inspectionStep);
  }

  private InspectionResult executeFullInspectionStep(
      InspectionStep inspectionStep, ExcelWorkbook workbook, WorkbookLocation workbookLocation) {
    SemanticSelectorResolver.ResolvedInspectionTarget resolvedTarget =
        selectorResolver.resolveInspectionTarget(
            inspectionStep.stepId(), workbook, inspectionStep.target(), inspectionStep.query());
    if (resolvedTarget.isShortCircuit()) {
      return Objects.requireNonNull(
          resolvedTarget.shortCircuitResult(), "short-circuit inspection result must be present");
    }
    dev.erst.gridgrind.excel.WorkbookReadResult result =
        workbookEngine
            .read(
                workbook,
                workbookLocation,
                InspectionCommandConverter.toReadCommand(
                    inspectionStep.stepId(),
                    Objects.requireNonNull(
                        resolvedTarget.selector(), "resolved selector must be present"),
                    inspectionStep.query()))
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
    if (!(command instanceof WorkbookReadCommand.Introspection introspection)) {
      throw new IllegalArgumentException(
          GridGrindExecutionModeMetadata.eventRead()
              .unsupportedQueryMessage(inspectionStep.query().queryType()));
    }
    dev.erst.gridgrind.excel.WorkbookReadResult result =
        eventWorkbookReader.apply(workbookPath, List.of(introspection)).getFirst();
    return InspectionResultConverter.toReadResult(result);
  }
}
