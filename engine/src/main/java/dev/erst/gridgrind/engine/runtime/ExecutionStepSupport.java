package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.catalog.GridGrindExecutionModeMetadata;
import dev.erst.gridgrind.core.grant.ExecutionGrant;
import dev.erst.gridgrind.core.plan.AssertionResult;
import dev.erst.gridgrind.core.plan.AssertionStep;
import dev.erst.gridgrind.core.plan.ExecutionModeSpec;
import dev.erst.gridgrind.core.plan.InspectionResult;
import dev.erst.gridgrind.core.plan.InspectionStep;
import dev.erst.gridgrind.excel.ExcelTempFileWriteTargetSupport;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.ExcelWorkbooks;
import dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition;
import dev.erst.gridgrind.excel.WorkbookExecutionEngine;
import dev.erst.gridgrind.excel.WorkbookLocation;
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
      TempFileFactory tempFileFactory,
      ExecutionInputBindings bindings) {
    this.workbookEngine = Objects.requireNonNull(workbookEngine, "workbookEngine must not be null");
    this.selectorResolver =
        Objects.requireNonNull(selectorResolver, "selectorResolver must not be null");
    this.assertionExecutor =
        Objects.requireNonNull(assertionExecutor, "assertionExecutor must not be null");
    this.tempFileFactory =
        Objects.requireNonNull(tempFileFactory, "tempFileFactory must not be null");
    this.mutationStepExecutor =
        new FormulaAwareMutationStepExecutor(this.selectorResolver, bindings);
  }

  ExecutionStepSupport(
      WorkbookExecutionEngine workbookEngine,
      SemanticSelectorResolver selectorResolver,
      AssertionExecutor assertionExecutor,
      TempFileFactory tempFileFactory) {
    this(
        workbookEngine,
        selectorResolver,
        assertionExecutor,
        tempFileFactory,
        new ExecutionInputBindings(
            Path.of("."),
            Path.of("tmp", "gridgrind-step-support"),
            new ExecutionGrant(
                List.of(),
                List.of(),
                new ExecutionGrant.WorkbookTargetAuthority.WorkbookWide(),
                new ExecutionGrant.PublicationAuthority.None(),
                List.of(),
                dev.erst.gridgrind.core.grant.AcceptancePolicy.minimum())));
  }

  FormulaAwareMutationStepExecutor mutationStepExecutor() {
    return mutationStepExecutor;
  }

  InspectionResult executeInspectionStep(
      InspectionStep inspectionStep,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation,
      ExecutionModeSpec executionMode)
      throws IOException {
    if (!(executionMode instanceof ExecutionModeSpec.EventRead)) {
      return executeFullInspectionStep(inspectionStep, workbook, workbookLocation);
    }
    return executeEventInspectionAgainstWorkbook(inspectionStep, workbook);
  }

  AssertionResult executeAssertionStep(
      AssertionStep assertionStep,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation,
      ExecutionModeSpec executionMode)
      throws IOException, AssertionFailedException {
    if (!(executionMode instanceof ExecutionModeSpec.EventRead)) {
      return assertionExecutor.execute(assertionStep, workbook, workbookLocation);
    }
    throw new IllegalStateException(
        "execution.mode.type=EVENT_READ does not support assertion steps");
  }

  AssertionStepExecution executeAssertionStepCollecting(
      AssertionStep assertionStep,
      ExcelWorkbook workbook,
      WorkbookLocation workbookLocation,
      ExecutionModeSpec executionMode)
      throws IOException {
    if (!(executionMode instanceof ExecutionModeSpec.EventRead)) {
      return assertionExecutor.executeCollecting(assertionStep, workbook, workbookLocation);
    }
    throw new IllegalStateException(
        "execution.mode.type=EVENT_READ does not support assertion steps");
  }

  InspectionResult executeInspectionAgainstMaterializedPath(
      InspectionStep inspectionStep,
      WorkbookLocation workbookLocation,
      ExecutionModeSpec executionMode,
      Path materializedPath)
      throws IOException {
    if (!(executionMode instanceof ExecutionModeSpec.EventRead)) {
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
    dev.erst.gridgrind.core.plan.InspectionResult result =
        workbookEngine
            .read(
                workbook,
                workbookLocation,
                resolvedInspection(
                    inspectionStep.stepId(),
                    Objects.requireNonNull(
                        resolvedTarget.selector(), "resolved selector must be present"),
                    inspectionStep.query()))
            .getFirst();
    return result;
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
      ExecutionModeSpec executionMode)
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

  AssertionResult executeAssertionAgainstMaterializedPath(
      AssertionStep assertionStep, WorkbookLocation workbookLocation, Path materializedPath)
      throws IOException, AssertionFailedException {
    return executeFullAssertionAgainstMaterializedPath(
        assertionStep, workbookLocation, materializedPath);
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
    dev.erst.gridgrind.core.plan.ResolvedInspectionStep command =
        resolvedInspection(inspectionStep);
    if (!(command.query() instanceof dev.erst.gridgrind.core.plan.InspectionQuery.Introspection)) {
      throw new IllegalArgumentException(
          GridGrindExecutionModeMetadata.eventRead()
              .unsupportedQueryMessage(inspectionStep.query().queryType()));
    }
    dev.erst.gridgrind.core.plan.InspectionResult result =
        eventWorkbookReader.apply(workbookPath, List.of(command)).getFirst();
    return result;
  }

  private static dev.erst.gridgrind.core.plan.ResolvedInspectionStep resolvedInspection(
      dev.erst.gridgrind.core.vocabulary.StepId id,
      dev.erst.gridgrind.core.plan.Selector target,
      dev.erst.gridgrind.core.plan.InspectionQuery query) {
    return new dev.erst.gridgrind.core.plan.ResolvedInspectionStep(
        id, (dev.erst.gridgrind.core.plan.ResolvedTarget) target, query);
  }

  private static dev.erst.gridgrind.core.plan.ResolvedInspectionStep resolvedInspection(
      dev.erst.gridgrind.core.plan.InspectionStep step) {
    return resolvedInspection(step.stepId(), step.target(), step.query());
  }
}
