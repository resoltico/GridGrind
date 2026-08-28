package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.contract.assertion.AssertionResult;
import dev.erst.gridgrind.contract.catalog.GridGrindExecutionModeMetadata;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.contract.step.AssertionStep;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.excel.ExcelTempFileWriteTargetSupport;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.ExcelWorkbooks;
import dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition;
import dev.erst.gridgrind.excel.WorkbookExecutionEngine;
import dev.erst.gridgrind.excel.WorkbookLocation;
import dev.erst.gridgrind.excel.WorkbookReadCommand;
import dev.erst.gridgrind.excel.event.ExcelEventWorkbookReader;
import dev.erst.gridgrind.excel.stream.ExcelStreamingWorkbookWriter;
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
  private final FormulaAwareMutationStepExecutor mutationStepExecutor;

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
    this.mutationStepExecutor =
        new FormulaAwareMutationStepExecutor(this.workbookEngine, this.selectorResolver);
  }

  FormulaAwareMutationStepExecutor mutationStepExecutor() {
    return mutationStepExecutor;
  }

  InspectionResult executeInspectionStep(
      InspectionStep inspectionStep,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation,
      ExecutionModeInput executionMode)
      throws IOException {
    if (!(executionMode instanceof ExecutionModeInput.EventRead)) {
      return executeFullInspectionStep(inspectionStep, workbook, workbookLocation);
    }
    return executeEventInspectionAgainstWorkbook(inspectionStep, workbook);
  }

  AssertionResult executeAssertionStep(
      AssertionStep assertionStep,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation,
      ExecutionModeInput executionMode)
      throws IOException, AssertionFailedException {
    if (!(executionMode instanceof ExecutionModeInput.EventRead)) {
      return assertionExecutor.execute(assertionStep, workbook, workbookLocation);
    }
    throw new IllegalStateException(
        "execution.mode.type=EVENT_READ does not support assertion steps");
  }

  AssertionStepExecution executeAssertionStepCollecting(
      AssertionStep assertionStep,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation,
      ExecutionModeInput executionMode)
      throws IOException {
    if (!(executionMode instanceof ExecutionModeInput.EventRead)) {
      return assertionExecutor.executeCollecting(assertionStep, workbook, workbookLocation);
    }
    throw new IllegalStateException(
        "execution.mode.type=EVENT_READ does not support assertion steps");
  }

  InspectionResult executeInspectionAgainstMaterializedPath(
      InspectionStep inspectionStep,
      WorkbookLocation workbookLocation,
      ExecutionModeInput executionMode,
      Path materializedPath)
      throws IOException {
    if (!(executionMode instanceof ExecutionModeInput.EventRead)) {
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
      tempPath =
          ExcelTempFileWriteTargetSupport.prepareCreateNewTarget(
              tempFileFactory.createTempFile("gridgrind-event-read-", ".xlsx"));
      workbook
          .persistence()
          .savePlainWorkbook(tempPath, WorkbookArtifactWriteDisposition.CREATE_NEW);
      return executeEventInspection(tempPath, inspectionStep);
    } finally {
      ExecutionWorkbookSupport.deleteIfExists(tempPath);
    }
  }

  InspectionResult executeStreamingInspectionStep(
      ExcelStreamingWorkbookWriter writer,
      InspectionStep inspectionStep,
      WorkbookLocation workbookLocation,
      ExecutionModeInput executionMode)
      throws IOException {
    Path tempPath =
        ExcelTempFileWriteTargetSupport.prepareCreateNewTarget(
            tempFileFactory.createTempFile("gridgrind-streaming-step-", ".xlsx"));
    try {
      writer.save(tempPath, WorkbookArtifactWriteDisposition.CREATE_NEW);
      return executeInspectionAgainstMaterializedPath(
          inspectionStep, workbookLocation, executionMode, tempPath);
    } finally {
      ExecutionWorkbookSupport.deleteIfExists(tempPath);
    }
  }

  AssertionResult executeStreamingAssertionStep(
      ExcelStreamingWorkbookWriter writer,
      AssertionStep assertionStep,
      WorkbookLocation workbookLocation)
      throws IOException, AssertionFailedException {
    Path tempPath =
        ExcelTempFileWriteTargetSupport.prepareCreateNewTarget(
            tempFileFactory.createTempFile("gridgrind-streaming-step-", ".xlsx"));
    try {
      writer.save(tempPath, WorkbookArtifactWriteDisposition.CREATE_NEW);
      return executeFullAssertionAgainstMaterializedPath(assertionStep, workbookLocation, tempPath);
    } finally {
      ExecutionWorkbookSupport.deleteIfExists(tempPath);
    }
  }

  AssertionStepExecution executeStreamingAssertionStepCollecting(
      ExcelStreamingWorkbookWriter writer,
      AssertionStep assertionStep,
      WorkbookLocation workbookLocation)
      throws IOException {
    Path tempPath =
        ExcelTempFileWriteTargetSupport.prepareCreateNewTarget(
            tempFileFactory.createTempFile("gridgrind-streaming-step-", ".xlsx"));
    try {
      writer.save(tempPath, WorkbookArtifactWriteDisposition.CREATE_NEW);
      return executeFullAssertionCollectingAgainstMaterializedPath(
          assertionStep, workbookLocation, tempPath);
    } finally {
      ExecutionWorkbookSupport.deleteIfExists(tempPath);
    }
  }

  private InspectionResult executeFullInspectionAgainstMaterializedPath(
      InspectionStep inspectionStep, WorkbookLocation workbookLocation, Path materializedPath)
      throws IOException {
    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(
            materializedPath,
            dev.erst.gridgrind.excel.ExcelFormulaEnvironment.defaults(),
            tempFileFactory::createTempFile)) {
      return executeFullInspectionStep(inspectionStep, workbook, workbookLocation);
    }
  }

  private AssertionResult executeFullAssertionAgainstMaterializedPath(
      AssertionStep assertionStep, WorkbookLocation workbookLocation, Path materializedPath)
      throws IOException, AssertionFailedException {
    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(
            materializedPath,
            dev.erst.gridgrind.excel.ExcelFormulaEnvironment.defaults(),
            tempFileFactory::createTempFile)) {
      return assertionExecutor.execute(assertionStep, workbook, workbookLocation);
    }
  }

  private AssertionStepExecution executeFullAssertionCollectingAgainstMaterializedPath(
      AssertionStep assertionStep, WorkbookLocation workbookLocation, Path materializedPath)
      throws IOException {
    try (ExcelWorkbook workbook =
        ExcelWorkbooks.open(
            materializedPath,
            dev.erst.gridgrind.excel.ExcelFormulaEnvironment.defaults(),
            tempFileFactory::createTempFile)) {
      return assertionExecutor.executeCollecting(assertionStep, workbook, workbookLocation);
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
