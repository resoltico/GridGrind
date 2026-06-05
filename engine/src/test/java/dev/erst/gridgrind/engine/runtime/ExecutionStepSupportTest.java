package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.catalog.GridGrindExecutionModeMetadata;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.query.*;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.excel.WorkbookExecutionEngine;
import dev.erst.gridgrind.excel.WorkbookLocation;
import dev.erst.gridgrind.excel.WorkbookTempFileFactory;
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
    ExecutionStepSupport support =
        new ExecutionStepSupport(
            workbookEngine,
            selectorResolver,
            new AssertionExecutor(workbookEngine, selectorResolver),
            (prefix, suffix) -> {
              throw new AssertionError("temp file creation should not happen for this branch");
            });
    InspectionStep inspectionStep =
        new InspectionStep(
            "formula-surface",
            new SheetSelector.All(),
            new InspectionSurfaceQuery.GetFormulaSurface());

    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () -> support.executeEventInspection(Path.of("unused.xlsx"), inspectionStep));

    assertEquals(
        GridGrindExecutionModeMetadata.eventRead()
            .unsupportedQueryMessage(inspectionStep.query().queryType()),
        failure.getMessage());
  }

  @Test
  void fullInspectionMaterializationAcceptsRootAnchorsWithoutAParentDirectory() throws IOException {
    WorkbookExecutionEngine workbookEngine = new WorkbookExecutionEngine();
    SemanticSelectorResolver selectorResolver = new SemanticSelectorResolver(workbookEngine);
    Path tempRoot = Files.createTempDirectory("gridgrind-step-support-root-");
    ExecutionStepSupport support =
        new ExecutionStepSupport(
            workbookEngine,
            selectorResolver,
            new AssertionExecutor(workbookEngine, selectorResolver),
            WorkbookTempFileFactory.rooted(tempRoot)::createTempFile);
    InspectionStep inspectionStep =
        new InspectionStep(
            "formula-surface",
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
                ExecutionModeInput.fullXssf(),
                filesystemRoot));
  }
}
