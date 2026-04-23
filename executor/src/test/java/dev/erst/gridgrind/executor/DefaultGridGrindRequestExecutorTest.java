package dev.erst.gridgrind.executor;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;

import org.junit.jupiter.api.Test;

/**
 * Incremental-build compatibility shim for the historical monolithic executor test class.
 *
 * <p>The executor test surface is now split by responsibility, but preserving this compiled class
 * name prevents stale no-clean builds from discovering a deleted test artifact.
 */
class DefaultGridGrindRequestExecutorTest {
  @Test
  void preservesWorkbookWorkflowCoverageEntryPoint() throws Exception {
    assertDoesNotThrow(
        () -> {
          DefaultGridGrindRequestExecutorWorkbookWorkflowTest tests =
              new DefaultGridGrindRequestExecutorWorkbookWorkflowTest();
          tests.executesAssertionStepsAlongsideMutationsAndInspections();
          tests.executesWorkbookWorkflowAndReturnsOrderedReadResults();
        });
  }

  @Test
  void preservesFailureAndPersistenceCoverageEntryPoint() throws Exception {
    assertDoesNotThrow(
        () -> {
          DefaultGridGrindRequestExecutorFailureAndPersistenceTest tests =
              new DefaultGridGrindRequestExecutorFailureAndPersistenceTest();
          tests.returnsStructuredFailureWithOperationContext();
          tests.preservesStructuralWorkbookStateAcrossExistingWorkbookRoundTrips();
        });
  }

  @Test
  void preservesStyleAndFormulaCoverageEntryPoint() throws Exception {
    assertDoesNotThrow(
        () -> {
          DefaultGridGrindRequestExecutorStyleAndFormulaTest tests =
              new DefaultGridGrindRequestExecutorStyleAndFormulaTest();
          tests.executesRangeAndStyleWorkflowAndSurfacesStyledCells();
          tests.executesFormattingDepthWorkflowAndPersistsReportedStyleState();
        });
  }

  @Test
  void preservesTranslationCoverageEntryPoint() {
    assertDoesNotThrow(
        () -> {
          DefaultGridGrindRequestExecutorTranslationTest tests =
              new DefaultGridGrindRequestExecutorTranslationTest();
          tests.convertsWaveThreeOperationsIntoWorkbookCommands();
          tests.convertsReadResultsIntoProtocolReadResults();
        });
  }
}
