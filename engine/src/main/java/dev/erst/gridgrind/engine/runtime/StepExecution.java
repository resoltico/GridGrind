package dev.erst.gridgrind.engine.runtime;

import dev.erst.gridgrind.core.plan.AssertionResult;
import dev.erst.gridgrind.core.plan.AssertionStep;
import dev.erst.gridgrind.core.plan.ExecutionModeSpec;
import dev.erst.gridgrind.core.plan.InspectionResult;
import dev.erst.gridgrind.core.plan.InspectionStep;
import dev.erst.gridgrind.core.semantics.BackendCapabilities;
import dev.erst.gridgrind.engine.backend.eventread.EventWorkbookReader;
import dev.erst.gridgrind.engine.backend.streaming.StreamingWorkbookWriter;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookArtifactWriteDisposition;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookExecutionEngine;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookLocation;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookTempFileFactory;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbook;
import dev.erst.gridgrind.engine.backend.xssf.XssfWorkbooks;
import java.io.IOException;
import java.nio.file.Path;
import java.util.List;
import java.util.Objects;

/** Executes workbook steps and materialized-read seams for the request executor. */
final class StepExecution {
  private final WorkbookExecutionEngine workbookEngine;
  private final SemanticSelectorResolver selectorResolver;
  private final AssertionExecutor assertionExecutor;
  private final TempFileFactory tempFileFactory;
  private final FormulaAwareMutationStepExecutor mutationStepExecutor;

  StepExecution(
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

  FormulaAwareMutationStepExecutor mutationStepExecutor() {
    return mutationStepExecutor;
  }

  InspectionResult executeInspectionStep(
      InspectionStep inspectionStep,
      XssfWorkbook workbook,
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
      XssfWorkbook workbook,
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
      XssfWorkbook workbook,
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
      InspectionStep inspectionStep, XssfWorkbook workbook, WorkbookLocation workbookLocation) {
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
      InspectionStep inspectionStep, XssfWorkbook workbook) throws IOException {
    Path tempPath = null;
    try {
      tempPath =
          WorkbookTempFileFactory.prepareCreateNewTarget(
              tempFileFactory.createTempFile("gridgrind-event-read-", ".xlsx"));
      workbook
          .persistence()
          .savePlainWorkbook(tempPath, WorkbookArtifactWriteDisposition.CREATE_NEW);
      return executeEventInspection(tempPath, inspectionStep);
    } finally {
      WorkbookLifecycle.deleteIfExists(tempPath);
    }
  }

  InspectionResult executeStreamingInspectionStep(
      StreamingWorkbookWriter writer,
      InspectionStep inspectionStep,
      WorkbookLocation workbookLocation,
      ExecutionModeSpec executionMode)
      throws IOException {
    Path tempPath =
        WorkbookTempFileFactory.prepareCreateNewTarget(
            tempFileFactory.createTempFile("gridgrind-streaming-step-", ".xlsx"));
    try {
      writer.save(tempPath, WorkbookArtifactWriteDisposition.CREATE_NEW);
      return executeInspectionAgainstMaterializedPath(
          inspectionStep, workbookLocation, executionMode, tempPath);
    } finally {
      WorkbookLifecycle.deleteIfExists(tempPath);
    }
  }

  AssertionResult executeStreamingAssertionStep(
      StreamingWorkbookWriter writer,
      AssertionStep assertionStep,
      WorkbookLocation workbookLocation)
      throws IOException, AssertionFailedException {
    Path tempPath =
        WorkbookTempFileFactory.prepareCreateNewTarget(
            tempFileFactory.createTempFile("gridgrind-streaming-step-", ".xlsx"));
    try {
      writer.save(tempPath, WorkbookArtifactWriteDisposition.CREATE_NEW);
      return executeFullAssertionAgainstMaterializedPath(assertionStep, workbookLocation, tempPath);
    } finally {
      WorkbookLifecycle.deleteIfExists(tempPath);
    }
  }

  AssertionResult executeAssertionAgainstMaterializedPath(
      AssertionStep assertionStep, WorkbookLocation workbookLocation, Path materializedPath)
      throws IOException, AssertionFailedException {
    return executeFullAssertionAgainstMaterializedPath(
        assertionStep, workbookLocation, materializedPath);
  }

  AssertionStepExecution executeStreamingAssertionStepCollecting(
      StreamingWorkbookWriter writer,
      AssertionStep assertionStep,
      WorkbookLocation workbookLocation)
      throws IOException {
    Path tempPath =
        WorkbookTempFileFactory.prepareCreateNewTarget(
            tempFileFactory.createTempFile("gridgrind-streaming-step-", ".xlsx"));
    try {
      writer.save(tempPath, WorkbookArtifactWriteDisposition.CREATE_NEW);
      return executeFullAssertionCollectingAgainstMaterializedPath(
          assertionStep, workbookLocation, tempPath);
    } finally {
      WorkbookLifecycle.deleteIfExists(tempPath);
    }
  }

  private InspectionResult executeFullInspectionAgainstMaterializedPath(
      InspectionStep inspectionStep, WorkbookLocation workbookLocation, Path materializedPath)
      throws IOException {
    try (XssfWorkbook workbook =
        XssfWorkbooks.open(
            materializedPath,
            dev.erst.gridgrind.core.model.formula.ResolvedFormulaEnvironment.defaults(),
            tempFileFactory::createTempFile)) {
      return executeFullInspectionStep(inspectionStep, workbook, workbookLocation);
    }
  }

  private AssertionResult executeFullAssertionAgainstMaterializedPath(
      AssertionStep assertionStep, WorkbookLocation workbookLocation, Path materializedPath)
      throws IOException, AssertionFailedException {
    try (XssfWorkbook workbook =
        XssfWorkbooks.open(
            materializedPath,
            dev.erst.gridgrind.core.model.formula.ResolvedFormulaEnvironment.defaults(),
            tempFileFactory::createTempFile)) {
      return assertionExecutor.execute(assertionStep, workbook, workbookLocation);
    }
  }

  private AssertionStepExecution executeFullAssertionCollectingAgainstMaterializedPath(
      AssertionStep assertionStep, WorkbookLocation workbookLocation, Path materializedPath)
      throws IOException {
    try (XssfWorkbook workbook =
        XssfWorkbooks.open(
            materializedPath,
            dev.erst.gridgrind.core.model.formula.ResolvedFormulaEnvironment.defaults(),
            tempFileFactory::createTempFile)) {
      return assertionExecutor.executeCollecting(assertionStep, workbook, workbookLocation);
    }
  }

  InspectionResult executeEventInspection(Path workbookPath, InspectionStep inspectionStep)
      throws IOException {
    EventWorkbookReader eventWorkbookReader = new EventWorkbookReader();
    dev.erst.gridgrind.core.plan.ResolvedInspectionStep command =
        resolvedInspection(inspectionStep);
    if (!(command.query() instanceof dev.erst.gridgrind.core.plan.InspectionQuery.Introspection)) {
      throw new IllegalArgumentException(
          BackendCapabilities.eventRead()
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
