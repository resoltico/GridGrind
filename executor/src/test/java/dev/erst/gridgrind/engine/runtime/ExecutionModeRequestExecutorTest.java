package dev.erst.gridgrind.engine.runtime;

import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.*;
import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.assertion.*;
import dev.erst.gridgrind.contract.assertion.ExpectedCellValue;
import dev.erst.gridgrind.contract.catalog.GridGrindContractText;
import dev.erst.gridgrind.contract.dto.ExecutionModeInput;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.GridGrindResponsePersistence;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.*;
import dev.erst.gridgrind.contract.query.SheetInspectionResult;
import dev.erst.gridgrind.contract.query.WorkbookInspectionResult;
import dev.erst.gridgrind.contract.selector.*;
import dev.erst.gridgrind.contract.step.InspectionStep;
import dev.erst.gridgrind.excel.ExcelCellValue;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookExecutionEngine;
import dev.erst.gridgrind.excel.foundation.ExcelSheetVisibility;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;

/** End-to-end tests for explicit low-memory execution-mode request workflows. */
class ExecutionModeRequestExecutorTest {
  @Test
  void eventReadModeMatchesFullXssfForExistingWorkbookSummaries() throws IOException {
    Path workbookPath = Files.createTempFile("gridgrind-event-mode-existing-", ".xlsx");
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Ops");
      workbook.getOrCreateSheet("Archive");
      workbook.setSheetVisibility("Archive", ExcelSheetVisibility.HIDDEN);
      workbook.setActiveSheet("Ops");
      workbook.setSelectedSheets(List.of("Ops"));
      workbook.formulas().markRecalculateOnOpen();
      workbook.sheet("Ops").setCell("A1", ExcelCellValue.text("Header"));
      workbook.sheet("Ops").setCell("C2", ExcelCellValue.number(42.0d));
      workbook.save(workbookPath);
    }

    List<InspectionStep> reads =
        List.of(
            inspect(
                "workbook",
                new WorkbookSelector.Current(),
                new WorkbookIntrospectionQuery.GetWorkbookSummary()),
            inspect(
                "sheet",
                new SheetSelector.ByName("Ops"),
                new SheetIntrospectionQuery.GetSheetSummary()));
    GridGrindResponse.Success full =
        success(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        List.of(),
                        reads)));
    GridGrindResponse.Success event =
        success(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        new ExecutionModeInput(
                            ExecutionModeInput.ReadMode.EVENT_READ,
                            ExecutionModeInput.WriteMode.FULL_XSSF),
                        null,
                        List.of(),
                        reads)));

    assertEquals(full.inspections(), event.inspections());
    assertEquals(full.persistence(), event.persistence());
    assertEquals(List.of(), event.warnings());
  }

  @Test
  void eventReadModeRejectsMutationWorkflowsUpFront() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        new ExecutionModeInput(
                            ExecutionModeInput.ReadMode.EVENT_READ,
                            ExecutionModeInput.WriteMode.FULL_XSSF),
                        null,
                        List.of(
                            mutate(
                                new SheetSelector.ByName("Ops"),
                                new WorkbookMutationAction.EnsureSheet())),
                        List.of(
                            inspect(
                                "workbook",
                                new WorkbookSelector.Current(),
                                new WorkbookIntrospectionQuery.GetWorkbookSummary())))));

    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertEquals("VALIDATE_REQUEST", failure.problem().context().stage());
    assertTrue(failure.problem().message().contains("inspection steps only"));
    assertTrue(failure.problem().message().contains("MUTATION"));
  }

  @Test
  void streamingWriteModePersistsWorkbookAndMatchesFullReadback() throws IOException {
    Path workbookPath = Files.createTempFile("gridgrind-streaming-mode-", ".xlsx");
    Files.deleteIfExists(workbookPath);

    List<ExecutorTestPlanSupport.PendingMutation> operations =
        List.of(
            mutate(new SheetSelector.ByName("Ops"), new WorkbookMutationAction.EnsureSheet()),
            mutate(
                new SheetSelector.ByName("Ops"),
                new CellMutationAction.AppendRow(List.of(textCell("Item"), textCell("Total")))),
            mutate(
                new SheetSelector.ByName("Ops"),
                new CellMutationAction.AppendRow(
                    List.of(textCell("Hosting"), formulaCell("2+3")))));
    List<InspectionStep> reads =
        List.of(
            inspect(
                "workbook",
                new WorkbookSelector.Current(),
                new WorkbookIntrospectionQuery.GetWorkbookSummary()),
            inspect(
                "cells",
                new CellSelector.ByAddresses("Ops", List.of("A1", "B2")),
                new SheetIntrospectionQuery.GetCells()));

    GridGrindResponse.Success streaming =
        success(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.SaveAs(workbookPath.toString()),
                        executionPolicy(
                            new ExecutionModeInput(
                                ExecutionModeInput.ReadMode.FULL_XSSF,
                                ExecutionModeInput.WriteMode.STREAMING_WRITE),
                            markRecalculateOnOpen()),
                        null,
                        List.copyOf(operations),
                        List.copyOf(reads))));
    GridGrindResponse.Success reopened =
        success(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        List.of(),
                        List.copyOf(reads))));

    assertTrue(Files.exists(workbookPath));
    WorkbookInspectionResult.WorkbookSummaryResult streamingWorkbookSummary =
        assertInstanceOf(
            WorkbookInspectionResult.WorkbookSummaryResult.class,
            streaming.inspections().getFirst());
    WorkbookInspectionResult.WorkbookSummaryResult reopenedWorkbookSummary =
        assertInstanceOf(
            WorkbookInspectionResult.WorkbookSummaryResult.class,
            reopened.inspections().getFirst());
    SheetInspectionResult.CellsResult streamingCells =
        assertInstanceOf(SheetInspectionResult.CellsResult.class, streaming.inspections().get(1));
    SheetInspectionResult.CellsResult reopenedCells =
        assertInstanceOf(SheetInspectionResult.CellsResult.class, reopened.inspections().get(1));

    assertFalse(streamingWorkbookSummary.workbook().forceFormulaRecalculationOnOpen());
    assertTrue(reopenedWorkbookSummary.workbook().forceFormulaRecalculationOnOpen());
    assertEquals(streamingCells, reopenedCells);
    assertEquals(workbookPath.toAbsolutePath().toString(), savedPath(streaming));
  }

  @Test
  void streamingWriteModeRejectsEventReadInspectionModeWhenMutationsExist() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        new ExecutionModeInput(
                            ExecutionModeInput.ReadMode.EVENT_READ,
                            ExecutionModeInput.WriteMode.STREAMING_WRITE),
                        null,
                        List.of(
                            mutate(
                                new SheetSelector.ByName("Ops"),
                                new WorkbookMutationAction.EnsureSheet())),
                        List.of(
                            inspect(
                                "workbook",
                                new WorkbookSelector.Current(),
                                new WorkbookIntrospectionQuery.GetWorkbookSummary())))));

    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertEquals("VALIDATE_REQUEST", failure.problem().context().stage());
    assertTrue(failure.problem().message().contains("inspection steps only"));
  }

  @Test
  void eventReadModeOnNewWorkbookWithoutMutationsUsesInjectedDefaultTempFactory() {
    DefaultGridGrindRequestExecutor executor =
        new DefaultGridGrindRequestExecutor(
            new DefaultGridGrindRequestExecutorDependencies(
                new WorkbookExecutionEngine(),
                ExcelWorkbook::close,
                java.nio.file.Files::createTempFile,
                dev.erst.gridgrind.excel.WorkbookArtifactIo.MaterializedWorkbook::close,
                dev.erst.gridgrind.excel.ExcelStreamingWorkbookWriter::markRecalculateOnOpen));

    GridGrindResponse.Success success =
        success(
            executor.execute(
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    new ExecutionModeInput(
                        ExecutionModeInput.ReadMode.EVENT_READ,
                        ExecutionModeInput.WriteMode.FULL_XSSF),
                    null,
                    List.of(),
                    List.of(
                        inspect(
                            "workbook",
                            new WorkbookSelector.Current(),
                            new WorkbookIntrospectionQuery.GetWorkbookSummary())))));

    WorkbookInspectionResult.WorkbookSummaryResult workbookSummary =
        assertInstanceOf(
            WorkbookInspectionResult.WorkbookSummaryResult.class, success.inspections().getFirst());
    assertEquals("workbook", workbookSummary.stepId());
    assertEquals(0, workbookSummary.workbook().sheetCount());
    assertEquals(List.of(), workbookSummary.workbook().sheetNames());
    assertInstanceOf(
        GridGrindResponsePersistence.PersistenceOutcome.NotSaved.class, success.persistence());
  }

  @Test
  void eventReadModeOnExistingWorkbookCanPersistCopyWhenDirectBypassIsIneligible()
      throws IOException {
    Path sourcePath = Files.createTempFile("gridgrind-event-read-saveas-source-", ".xlsx");
    Path persistedCopy = Files.createTempFile("gridgrind-event-read-saveas-copy-", ".xlsx");
    Files.deleteIfExists(persistedCopy);
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Ops");
      workbook.sheet("Ops").setCell("A1", ExcelCellValue.text("Header"));
      workbook.save(sourcePath);
    }

    GridGrindResponse.Success success =
        success(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.ExistingFile(sourcePath.toString()),
                        new WorkbookPlan.WorkbookPersistence.SaveAs(persistedCopy.toString()),
                        new ExecutionModeInput(
                            ExecutionModeInput.ReadMode.EVENT_READ,
                            ExecutionModeInput.WriteMode.FULL_XSSF),
                        null,
                        List.of(),
                        List.of(
                            inspect(
                                "workbook",
                                new WorkbookSelector.Current(),
                                new WorkbookIntrospectionQuery.GetWorkbookSummary())))));

    WorkbookInspectionResult.WorkbookSummaryResult workbookSummary =
        assertInstanceOf(
            WorkbookInspectionResult.WorkbookSummaryResult.class, success.inspections().getFirst());
    GridGrindResponsePersistence.PersistenceOutcome.SavedAs savedAs =
        assertInstanceOf(
            GridGrindResponsePersistence.PersistenceOutcome.SavedAs.class, success.persistence());

    assertEquals(1, workbookSummary.workbook().sheetCount());
    assertTrue(Files.exists(persistedCopy));
    assertEquals(persistedCopy.toAbsolutePath().toString(), savedAs.executionPath());
  }

  @Test
  void eventReadModeReturnsStructuredFailureForMissingExistingWorkbook() throws IOException {
    Path missingWorkbook =
        Path.of(System.getProperty("java.io.tmpdir"), "gridgrind-missing-event-read.xlsx");
    Files.deleteIfExists(missingWorkbook);

    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.ExistingFile(missingWorkbook.toString()),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        new ExecutionModeInput(
                            ExecutionModeInput.ReadMode.EVENT_READ,
                            ExecutionModeInput.WriteMode.FULL_XSSF),
                        null,
                        List.of(),
                        List.of(
                            inspect(
                                "workbook",
                                new WorkbookSelector.Current(),
                                new WorkbookIntrospectionQuery.GetWorkbookSummary())))));

    assertEquals(GridGrindProblemCode.WORKBOOK_NOT_FOUND, failure.problem().code());
    assertEquals("OPEN_WORKBOOK", failure.problem().context().stage());
  }

  @Test
  void eventReadModeSupportsEmptyReadListsWithoutMaterializingTempWorkbook() {
    GridGrindResponse.Success success =
        success(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        new ExecutionModeInput(
                            ExecutionModeInput.ReadMode.EVENT_READ,
                            ExecutionModeInput.WriteMode.FULL_XSSF),
                        null,
                        List.of(),
                        List.of())));

    assertEquals(List.of(), success.inspections());
    assertInstanceOf(
        GridGrindResponsePersistence.PersistenceOutcome.NotSaved.class, success.persistence());
  }

  @Test
  void eventReadModeRejectsAssertionStepsUpFront() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        new ExecutionModeInput(
                            ExecutionModeInput.ReadMode.EVENT_READ,
                            ExecutionModeInput.WriteMode.FULL_XSSF),
                        null,
                        List.of(),
                        List.of(
                            assertThat(
                                "assert-owner",
                                new CellSelector.ByAddress("Ops", "A1"),
                                new CellAssertion.CellValue(new ExpectedCellValue.Text("Owner")))),
                        List.of(
                            inspect(
                                "workbook",
                                new WorkbookSelector.Current(),
                                new WorkbookIntrospectionQuery.GetWorkbookSummary())))));

    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertEquals("VALIDATE_REQUEST", failure.problem().context().stage());
    assertTrue(failure.problem().message().contains("inspection steps only"));
    assertTrue(failure.problem().message().contains("ASSERTION"));
  }

  @Test
  void streamingWriteModeAllowsAppendRowValidationButFailsIfTheSheetWasNeverCreated() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        new ExecutionModeInput(
                            ExecutionModeInput.ReadMode.FULL_XSSF,
                            ExecutionModeInput.WriteMode.STREAMING_WRITE),
                        null,
                        List.of(
                            mutate(
                                new SheetSelector.ByName("Ops"),
                                new CellMutationAction.AppendRow(List.of(textCell("x"))))),
                        List.of())));

    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertEquals("VALIDATE_REQUEST", failure.problem().context().stage());
    assertTrue(failure.problem().message().contains("requires ENSURE_SHEET before APPEND_ROW"));
  }

  @Test
  void streamingWriteModeReportsTempFileCreationIoFailure() {
    DefaultGridGrindRequestExecutor executor =
        new DefaultGridGrindRequestExecutor(
            new DefaultGridGrindRequestExecutorDependencies(
                new WorkbookExecutionEngine(),
                ExcelWorkbook::close,
                (_, _) -> {
                  throw new IOException("temp creation failed");
                },
                dev.erst.gridgrind.excel.WorkbookArtifactIo.MaterializedWorkbook::close,
                dev.erst.gridgrind.excel.ExcelStreamingWorkbookWriter::markRecalculateOnOpen));

    GridGrindResponse.Failure failure =
        failure(
            executor.execute(
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    new ExecutionModeInput(
                        ExecutionModeInput.ReadMode.FULL_XSSF,
                        ExecutionModeInput.WriteMode.STREAMING_WRITE),
                    null,
                    List.of(
                        mutate(
                            new SheetSelector.ByName("Ops"),
                            new WorkbookMutationAction.EnsureSheet())),
                    List.of())));

    assertEquals(GridGrindProblemCode.IO_ERROR, failure.problem().code());
    assertEquals("EXECUTE_REQUEST", failure.problem().context().stage());
  }

  @Test
  void streamingWriteModeReturnsReadFailureAfterMaterialization() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        new ExecutionModeInput(
                            ExecutionModeInput.ReadMode.FULL_XSSF,
                            ExecutionModeInput.WriteMode.STREAMING_WRITE),
                        null,
                        List.of(
                            mutate(
                                new SheetSelector.ByName("Ops"),
                                new WorkbookMutationAction.EnsureSheet())),
                        List.of(
                            inspect(
                                "cells",
                                new CellSelector.ByAddresses("Missing", List.of("A1")),
                                new SheetIntrospectionQuery.GetCells())))));

    assertEquals(GridGrindProblemCode.SHEET_NOT_FOUND, failure.problem().code());
    assertEquals("EXECUTE_STEP", failure.problem().context().stage());
    assertEquals(
        java.util.Optional.of("Missing"),
        DefaultGridGrindRequestExecutorTestSupport.executeStepContext(failure).sheetName());
  }

  @Test
  void streamingWriteModeReturnsPersistFailureForInvalidSaveAsTarget() throws IOException {
    Path parentFile = Files.createTempFile("gridgrind-streaming-invalid-target-", ".tmp");
    Path workbookPath = parentFile.resolve("book.xlsx");

    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.SaveAs(workbookPath.toString()),
                        new ExecutionModeInput(
                            ExecutionModeInput.ReadMode.FULL_XSSF,
                            ExecutionModeInput.WriteMode.STREAMING_WRITE),
                        null,
                        List.of(
                            mutate(
                                new SheetSelector.ByName("Ops"),
                                new WorkbookMutationAction.EnsureSheet())),
                        List.of())));

    assertEquals(GridGrindProblemCode.IO_ERROR, failure.problem().code());
    assertEquals("PERSIST_WORKBOOK", failure.problem().context().stage());
    assertEquals(
        java.util.Optional.of(workbookPath.toAbsolutePath().toString()),
        DefaultGridGrindRequestExecutorTestSupport.persistWorkbookContext(failure)
            .persistencePath());
  }

  @Test
  void validatesUnsupportedExecutionModeCombinationsUpFront() throws IOException {
    Path workbookPath = Files.createTempFile("gridgrind-execution-mode-validation-", ".xlsx");
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Ops");
      workbook.save(workbookPath);
    }

    GridGrindResponse.Failure unsupportedEventRead =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        new ExecutionModeInput(
                            ExecutionModeInput.ReadMode.EVENT_READ,
                            ExecutionModeInput.WriteMode.FULL_XSSF),
                        null,
                        List.of(),
                        List.of(
                            inspect(
                                "cells",
                                new CellSelector.ByAddresses("Ops", List.of("A1")),
                                new SheetIntrospectionQuery.GetCells())))));
    assertEquals(GridGrindProblemCode.INVALID_REQUEST, unsupportedEventRead.problem().code());
    assertTrue(unsupportedEventRead.problem().message().contains("GET_CELLS"));

    GridGrindResponse.Failure existingStreamingSource =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.ExistingFile(workbookPath.toString()),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        new ExecutionModeInput(
                            ExecutionModeInput.ReadMode.FULL_XSSF,
                            ExecutionModeInput.WriteMode.STREAMING_WRITE),
                        null,
                        mutations(
                            mutate(
                                new SheetSelector.ByName("Ops"),
                                new CellMutationAction.AppendRow(List.of(textCell("x"))))),
                        inspections())));
    assertEquals(GridGrindProblemCode.INVALID_REQUEST, existingStreamingSource.problem().code());
    assertTrue(existingStreamingSource.problem().message().contains("requires source.type=NEW"));

    GridGrindResponse.Failure unsupportedStreamingOperation =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        new ExecutionModeInput(
                            ExecutionModeInput.ReadMode.FULL_XSSF,
                            ExecutionModeInput.WriteMode.STREAMING_WRITE),
                        null,
                        List.of(
                            mutate(
                                new SheetSelector.ByName("Ops"),
                                new WorkbookMutationAction.EnsureSheet()),
                            mutate(
                                new CellSelector.ByAddress("Ops", "A1"),
                                new CellMutationAction.SetCell(textCell("x")))),
                        inspections())));
    assertEquals(
        GridGrindProblemCode.INVALID_REQUEST, unsupportedStreamingOperation.problem().code());
    assertTrue(unsupportedStreamingOperation.problem().message().contains("SET_CELL"));
    assertTrue(
        unsupportedStreamingOperation
            .problem()
            .message()
            .contains(GridGrindContractText.streamingWriteMutationActionTypePhrase()));
    assertFalse(
        unsupportedStreamingOperation.problem().message().contains("FORCE_FORMULA_RECALC_ON_OPEN"));

    GridGrindResponse.Failure missingStreamingSheetMaterialization =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        executionPolicy(
                            new ExecutionModeInput(
                                ExecutionModeInput.ReadMode.FULL_XSSF,
                                ExecutionModeInput.WriteMode.STREAMING_WRITE),
                            markRecalculateOnOpen()),
                        null,
                        List.of(),
                        List.of())));
    assertEquals(
        GridGrindProblemCode.INVALID_REQUEST,
        missingStreamingSheetMaterialization.problem().code());
    assertTrue(
        missingStreamingSheetMaterialization
            .problem()
            .message()
            .contains("requires at least one ENSURE_SHEET mutation"));
  }

  private static GridGrindResponse.Success success(GridGrindResponse response) {
    return assertInstanceOf(GridGrindResponse.Success.class, response);
  }

  private static GridGrindResponse.Failure failure(GridGrindResponse response) {
    return assertInstanceOf(GridGrindResponse.Failure.class, response);
  }

  private static String savedPath(GridGrindResponse.Success success) {
    return switch (success.persistence()) {
      case GridGrindResponsePersistence.PersistenceOutcome.SavedAs savedAs ->
          savedAs.executionPath();
      case GridGrindResponsePersistence.PersistenceOutcome.Overwritten overwritten ->
          overwritten.executionPath();
      case GridGrindResponsePersistence.PersistenceOutcome.NotSaved _ ->
          throw new AssertionError("expected persisted workbook");
    };
  }
}
