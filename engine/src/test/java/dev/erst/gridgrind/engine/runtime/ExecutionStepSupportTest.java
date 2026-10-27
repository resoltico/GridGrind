package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.plan.*;
import dev.erst.gridgrind.core.plan.ExecutionModeSpec;
import dev.erst.gridgrind.core.plan.InspectionStep;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.semantics.BackendCapabilities;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookExecutionEngine;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookLocation;
import dev.erst.gridgrind.engine.backend.xssf.WorkbookTempFileFactory;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import org.junit.jupiter.api.Test;

/** Coverage for executor-owned step helpers inside the engine runtime. */
class ExecutionStepSupportTest {
  @Test
  void eventReadRejectsSurfaceQueriesBeforeTouchingTheWorkbookPath() {
    WorkbookExecutionEngine workbookEngine = new WorkbookExecutionEngine();
    SemanticSelectorResolver selectorResolver = new SemanticSelectorResolver(workbookEngine);
    StepExecution support =
        new StepExecution(
            workbookEngine,
            selectorResolver,
            new AssertionExecutor(workbookEngine, selectorResolver),
            (prefix, suffix) -> {
              throw new AssertionError("temp file creation should not happen for this branch");
            },
            new ExecutionInputBindings(
                Path.of("."),
                Path.of("tmp", "step-support"),
                ExecutionGrantTestSupport.noPublication()));
    InspectionStep inspectionStep =
        new InspectionStep(
            new StepId("formula-surface"),
            new SheetSelector.All(),
            new InspectionSurfaceQuery.GetFormulaSurface());

    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () -> support.executeEventInspection(Path.of("unused.xlsx"), inspectionStep));

    assertEquals(
        BackendCapabilities.eventRead().unsupportedQueryMessage(inspectionStep.query().queryType()),
        failure.getMessage());
  }

  @Test
  void fullInspectionMaterializationAcceptsRootAnchorsWithoutAParentDirectory() throws IOException {
    WorkbookExecutionEngine workbookEngine = new WorkbookExecutionEngine();
    SemanticSelectorResolver selectorResolver = new SemanticSelectorResolver(workbookEngine);
    Path tempRoot = Files.createTempDirectory("gridgrind-step-support-root-");
    StepExecution support =
        new StepExecution(
            workbookEngine,
            selectorResolver,
            new AssertionExecutor(workbookEngine, selectorResolver),
            WorkbookTempFileFactory.rooted(tempRoot)::createTempFile,
            new ExecutionInputBindings(
                tempRoot, tempRoot.resolve("private"), ExecutionGrantTestSupport.noPublication()));
    InspectionStep inspectionStep =
        new InspectionStep(
            new StepId("formula-surface"),
            new SheetSelector.All(),
            new InspectionSurfaceQuery.GetFormulaSurface());
    Path filesystemRoot = Path.of("").toAbsolutePath().normalize().getRoot();
    assertNotNull(filesystemRoot);

    assertThrows(
        IOException.class,
        () ->
            support.executeInspectionAgainstMaterializedPath(
                inspectionStep,
                new WorkbookLocation.UnsavedWorkbook(),
                ExecutionModeSpec.fullXssf(),
                filesystemRoot));
  }
}
