package dev.erst.gridgrind.protocol.exec;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.ExcelCellValue;
import dev.erst.gridgrind.excel.ExcelSheetVisibility;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.WorkbookCommandExecutor;
import dev.erst.gridgrind.excel.WorkbookReadExecutor;
import dev.erst.gridgrind.protocol.dto.ExecutionModeInput;
import dev.erst.gridgrind.protocol.dto.GridGrindProblemCode;
import dev.erst.gridgrind.protocol.dto.GridGrindRequest;
import dev.erst.gridgrind.protocol.dto.GridGrindResponse;
import dev.erst.gridgrind.protocol.operation.WorkbookOperation;
import dev.erst.gridgrind.protocol.read.WorkbookReadOperation;
import dev.erst.gridgrind.protocol.read.WorkbookReadResult;
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
      workbook.forceFormulaRecalculationOnOpen();
      workbook.sheet("Ops").setCell("A1", ExcelCellValue.text("Header"));
      workbook.sheet("Ops").setCell("C2", ExcelCellValue.number(42.0d));
      workbook.save(workbookPath);
    }

    List<WorkbookReadOperation> reads =
        List.of(
            new WorkbookReadOperation.GetWorkbookSummary("workbook"),
            new WorkbookReadOperation.GetSheetSummary("sheet", "Ops"));
    GridGrindResponse.Success full =
        success(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    new GridGrindRequest(
                        new GridGrindRequest.WorkbookSource.ExistingFile(workbookPath.toString()),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        List.of(),
                        reads)));
    GridGrindResponse.Success event =
        success(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    new GridGrindRequest(
                        new GridGrindRequest.WorkbookSource.ExistingFile(workbookPath.toString()),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        new ExecutionModeInput(
                            ExecutionModeInput.ReadMode.EVENT_READ,
                            ExecutionModeInput.WriteMode.FULL_XSSF),
                        null,
                        List.of(),
                        reads)));

    assertEquals(full.reads(), event.reads());
    assertEquals(full.persistence(), event.persistence());
    assertEquals(List.of(), event.warnings());
  }

  @Test
  void eventReadModeMatchesFullXssfAfterInMemoryMutations() {
    List<WorkbookOperation> operations =
        List.of(
            new WorkbookOperation.EnsureSheet("Ops"),
            new WorkbookOperation.AppendRow(
                "Ops", List.of(new dev.erst.gridgrind.protocol.dto.CellInput.Text("Header"))),
            new WorkbookOperation.AppendRow(
                "Ops", List.of(new dev.erst.gridgrind.protocol.dto.CellInput.Numeric(42.0d))));
    List<WorkbookReadOperation> reads =
        List.of(
            new WorkbookReadOperation.GetWorkbookSummary("workbook"),
            new WorkbookReadOperation.GetSheetSummary("sheet", "Ops"));

    GridGrindResponse.Success full =
        success(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    new GridGrindRequest(
                        new GridGrindRequest.WorkbookSource.New(),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        List.copyOf(operations),
                        List.copyOf(reads))));
    GridGrindResponse.Success event =
        success(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    new GridGrindRequest(
                        new GridGrindRequest.WorkbookSource.New(),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        new ExecutionModeInput(
                            ExecutionModeInput.ReadMode.EVENT_READ,
                            ExecutionModeInput.WriteMode.FULL_XSSF),
                        null,
                        List.copyOf(operations),
                        List.copyOf(reads))));

    assertEquals(full.reads(), event.reads());
    assertEquals(full.persistence(), event.persistence());
  }

  @Test
  void streamingWriteModePersistsWorkbookAndMatchesFullReadback() throws IOException {
    Path workbookPath = Files.createTempFile("gridgrind-streaming-mode-", ".xlsx");
    Files.deleteIfExists(workbookPath);

    List<WorkbookOperation> operations =
        List.of(
            new WorkbookOperation.EnsureSheet("Ops"),
            new WorkbookOperation.AppendRow(
                "Ops",
                List.of(
                    new dev.erst.gridgrind.protocol.dto.CellInput.Text("Item"),
                    new dev.erst.gridgrind.protocol.dto.CellInput.Text("Total"))),
            new WorkbookOperation.AppendRow(
                "Ops",
                List.of(
                    new dev.erst.gridgrind.protocol.dto.CellInput.Text("Hosting"),
                    new dev.erst.gridgrind.protocol.dto.CellInput.Formula("2+3"))),
            new WorkbookOperation.ForceFormulaRecalculationOnOpen());
    List<WorkbookReadOperation> reads =
        List.of(
            new WorkbookReadOperation.GetWorkbookSummary("workbook"),
            new WorkbookReadOperation.GetCells("cells", "Ops", List.of("A1", "B2")));

    GridGrindResponse.Success streaming =
        success(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    new GridGrindRequest(
                        new GridGrindRequest.WorkbookSource.New(),
                        new GridGrindRequest.WorkbookPersistence.SaveAs(workbookPath.toString()),
                        new ExecutionModeInput(
                            ExecutionModeInput.ReadMode.FULL_XSSF,
                            ExecutionModeInput.WriteMode.STREAMING_WRITE),
                        null,
                        List.copyOf(operations),
                        List.copyOf(reads))));
    GridGrindResponse.Success reopened =
        success(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    new GridGrindRequest(
                        new GridGrindRequest.WorkbookSource.ExistingFile(workbookPath.toString()),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        List.of(),
                        List.copyOf(reads))));

    assertTrue(Files.exists(workbookPath));
    assertEquals(reopened.reads(), streaming.reads());
    assertEquals(workbookPath.toAbsolutePath().toString(), savedPath(streaming));
  }

  @Test
  void streamingWriteModeCanUseEventReadSummaries() {
    List<WorkbookOperation> operations =
        List.of(
            new WorkbookOperation.EnsureSheet("Ops"),
            new WorkbookOperation.AppendRow(
                "Ops", List.of(new dev.erst.gridgrind.protocol.dto.CellInput.Text("Header"))));
    List<WorkbookReadOperation> reads =
        List.of(
            new WorkbookReadOperation.GetWorkbookSummary("workbook"),
            new WorkbookReadOperation.GetSheetSummary("sheet", "Ops"));

    GridGrindResponse.Success full =
        success(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    new GridGrindRequest(
                        new GridGrindRequest.WorkbookSource.New(),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        List.copyOf(operations),
                        List.copyOf(reads))));
    GridGrindResponse.Success lowMemory =
        success(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    new GridGrindRequest(
                        new GridGrindRequest.WorkbookSource.New(),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        new ExecutionModeInput(
                            ExecutionModeInput.ReadMode.EVENT_READ,
                            ExecutionModeInput.WriteMode.STREAMING_WRITE),
                        null,
                        List.copyOf(operations),
                        List.copyOf(reads))));

    assertEquals(full.reads(), lowMemory.reads());
    assertInstanceOf(GridGrindResponse.PersistenceOutcome.NotSaved.class, lowMemory.persistence());
  }

  @Test
  void eventReadModeOnNewWorkbookWithoutMutationsUsesInjectedDefaultTempFactory() {
    DefaultGridGrindRequestExecutor executor =
        new DefaultGridGrindRequestExecutor(
            new WorkbookCommandExecutor(), new WorkbookReadExecutor(), ExcelWorkbook::close);

    GridGrindResponse.Success success =
        success(
            executor.execute(
                new GridGrindRequest(
                    new GridGrindRequest.WorkbookSource.New(),
                    new GridGrindRequest.WorkbookPersistence.None(),
                    new ExecutionModeInput(
                        ExecutionModeInput.ReadMode.EVENT_READ,
                        ExecutionModeInput.WriteMode.FULL_XSSF),
                    null,
                    List.of(),
                    List.of(new WorkbookReadOperation.GetWorkbookSummary("workbook")))));

    WorkbookReadResult.WorkbookSummaryResult workbookSummary =
        assertInstanceOf(
            WorkbookReadResult.WorkbookSummaryResult.class, success.reads().getFirst());
    assertEquals("workbook", workbookSummary.requestId());
    assertEquals(0, workbookSummary.workbook().sheetCount());
    assertEquals(List.of(), workbookSummary.workbook().sheetNames());
    assertInstanceOf(GridGrindResponse.PersistenceOutcome.NotSaved.class, success.persistence());
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
                    new GridGrindRequest(
                        new GridGrindRequest.WorkbookSource.ExistingFile(sourcePath.toString()),
                        new GridGrindRequest.WorkbookPersistence.SaveAs(persistedCopy.toString()),
                        new ExecutionModeInput(
                            ExecutionModeInput.ReadMode.EVENT_READ,
                            ExecutionModeInput.WriteMode.FULL_XSSF),
                        null,
                        List.of(),
                        List.of(new WorkbookReadOperation.GetWorkbookSummary("workbook")))));

    WorkbookReadResult.WorkbookSummaryResult workbookSummary =
        assertInstanceOf(
            WorkbookReadResult.WorkbookSummaryResult.class, success.reads().getFirst());
    GridGrindResponse.PersistenceOutcome.SavedAs savedAs =
        assertInstanceOf(GridGrindResponse.PersistenceOutcome.SavedAs.class, success.persistence());

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
                    new GridGrindRequest(
                        new GridGrindRequest.WorkbookSource.ExistingFile(
                            missingWorkbook.toString()),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        new ExecutionModeInput(
                            ExecutionModeInput.ReadMode.EVENT_READ,
                            ExecutionModeInput.WriteMode.FULL_XSSF),
                        null,
                        List.of(),
                        List.of(new WorkbookReadOperation.GetWorkbookSummary("workbook")))));

    assertEquals(GridGrindProblemCode.IO_ERROR, failure.problem().code());
    assertEquals("EXECUTE_READ", failure.problem().context().stage());
  }

  @Test
  void eventReadModeSupportsEmptyReadListsWithoutMaterializingTempWorkbook() {
    GridGrindResponse.Success success =
        success(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    new GridGrindRequest(
                        new GridGrindRequest.WorkbookSource.New(),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        new ExecutionModeInput(
                            ExecutionModeInput.ReadMode.EVENT_READ,
                            ExecutionModeInput.WriteMode.FULL_XSSF),
                        null,
                        List.of(),
                        List.of())));

    assertEquals(List.of(), success.reads());
    assertInstanceOf(GridGrindResponse.PersistenceOutcome.NotSaved.class, success.persistence());
  }

  @Test
  void eventReadModeReportsTempFileCreationFailureWhenReadingMutatedWorkbook() {
    DefaultGridGrindRequestExecutor executor =
        new DefaultGridGrindRequestExecutor(
            new WorkbookCommandExecutor(),
            new WorkbookReadExecutor(),
            ExcelWorkbook::close,
            (_, _) -> {
              throw new IOException("temp creation failed");
            });

    GridGrindResponse.Failure failure =
        failure(
            executor.execute(
                new GridGrindRequest(
                    new GridGrindRequest.WorkbookSource.New(),
                    new GridGrindRequest.WorkbookPersistence.None(),
                    new ExecutionModeInput(
                        ExecutionModeInput.ReadMode.EVENT_READ,
                        ExecutionModeInput.WriteMode.FULL_XSSF),
                    null,
                    List.of(new WorkbookOperation.EnsureSheet("Ops")),
                    List.of(new WorkbookReadOperation.GetWorkbookSummary("workbook")))));

    assertEquals(GridGrindProblemCode.IO_ERROR, failure.problem().code());
    assertEquals("EXECUTE_READ", failure.problem().context().stage());
  }

  @Test
  void streamingWriteModeAllowsAppendRowValidationButFailsIfTheSheetWasNeverCreated() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    new GridGrindRequest(
                        new GridGrindRequest.WorkbookSource.New(),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        new ExecutionModeInput(
                            ExecutionModeInput.ReadMode.FULL_XSSF,
                            ExecutionModeInput.WriteMode.STREAMING_WRITE),
                        null,
                        List.of(
                            new WorkbookOperation.AppendRow(
                                "Ops",
                                List.of(new dev.erst.gridgrind.protocol.dto.CellInput.Text("x")))),
                        List.of())));

    assertEquals(GridGrindProblemCode.SHEET_NOT_FOUND, failure.problem().code());
    assertEquals("APPLY_OPERATION", failure.problem().context().stage());
    assertEquals("APPEND_ROW", failure.problem().context().operationType());
    assertEquals("Ops", failure.problem().context().sheetName());
  }

  @Test
  void streamingWriteModeReportsTempFileCreationIoFailure() {
    DefaultGridGrindRequestExecutor executor =
        new DefaultGridGrindRequestExecutor(
            new WorkbookCommandExecutor(),
            new WorkbookReadExecutor(),
            ExcelWorkbook::close,
            (_, _) -> {
              throw new IOException("temp creation failed");
            });

    GridGrindResponse.Failure failure =
        failure(
            executor.execute(
                new GridGrindRequest(
                    new GridGrindRequest.WorkbookSource.New(),
                    new GridGrindRequest.WorkbookPersistence.None(),
                    new ExecutionModeInput(
                        ExecutionModeInput.ReadMode.FULL_XSSF,
                        ExecutionModeInput.WriteMode.STREAMING_WRITE),
                    null,
                    List.of(new WorkbookOperation.EnsureSheet("Ops")),
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
                    new GridGrindRequest(
                        new GridGrindRequest.WorkbookSource.New(),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        new ExecutionModeInput(
                            ExecutionModeInput.ReadMode.FULL_XSSF,
                            ExecutionModeInput.WriteMode.STREAMING_WRITE),
                        null,
                        List.of(new WorkbookOperation.EnsureSheet("Ops")),
                        List.of(
                            new WorkbookReadOperation.GetCells(
                                "cells", "Missing", List.of("A1"))))));

    assertEquals(GridGrindProblemCode.SHEET_NOT_FOUND, failure.problem().code());
    assertEquals("EXECUTE_READ", failure.problem().context().stage());
    assertEquals("Missing", failure.problem().context().sheetName());
  }

  @Test
  void streamingWriteModeReturnsPersistFailureForInvalidSaveAsTarget() throws IOException {
    Path parentFile = Files.createTempFile("gridgrind-streaming-invalid-target-", ".tmp");
    Path workbookPath = parentFile.resolve("book.xlsx");

    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    new GridGrindRequest(
                        new GridGrindRequest.WorkbookSource.New(),
                        new GridGrindRequest.WorkbookPersistence.SaveAs(workbookPath.toString()),
                        new ExecutionModeInput(
                            ExecutionModeInput.ReadMode.FULL_XSSF,
                            ExecutionModeInput.WriteMode.STREAMING_WRITE),
                        null,
                        List.of(new WorkbookOperation.EnsureSheet("Ops")),
                        List.of())));

    assertEquals(GridGrindProblemCode.IO_ERROR, failure.problem().code());
    assertEquals("PERSIST_WORKBOOK", failure.problem().context().stage());
    assertEquals(
        workbookPath.toAbsolutePath().toString(), failure.problem().context().persistencePath());
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
                    new GridGrindRequest(
                        new GridGrindRequest.WorkbookSource.ExistingFile(workbookPath.toString()),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        new ExecutionModeInput(
                            ExecutionModeInput.ReadMode.EVENT_READ,
                            ExecutionModeInput.WriteMode.FULL_XSSF),
                        null,
                        List.of(),
                        List.of(
                            new WorkbookReadOperation.GetCells("cells", "Ops", List.of("A1"))))));
    assertEquals(GridGrindProblemCode.INVALID_REQUEST, unsupportedEventRead.problem().code());
    assertTrue(unsupportedEventRead.problem().message().contains("GET_CELLS"));

    GridGrindResponse.Failure existingStreamingSource =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    new GridGrindRequest(
                        new GridGrindRequest.WorkbookSource.ExistingFile(workbookPath.toString()),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        new ExecutionModeInput(
                            ExecutionModeInput.ReadMode.FULL_XSSF,
                            ExecutionModeInput.WriteMode.STREAMING_WRITE),
                        null,
                        List.of(
                            new WorkbookOperation.AppendRow(
                                "Ops",
                                List.of(new dev.erst.gridgrind.protocol.dto.CellInput.Text("x")))),
                        List.of())));
    assertEquals(GridGrindProblemCode.INVALID_REQUEST, existingStreamingSource.problem().code());
    assertTrue(existingStreamingSource.problem().message().contains("requires source.type=NEW"));

    GridGrindResponse.Failure unsupportedStreamingOperation =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    new GridGrindRequest(
                        new GridGrindRequest.WorkbookSource.New(),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        new ExecutionModeInput(
                            ExecutionModeInput.ReadMode.FULL_XSSF,
                            ExecutionModeInput.WriteMode.STREAMING_WRITE),
                        null,
                        List.of(
                            new WorkbookOperation.EnsureSheet("Ops"),
                            new WorkbookOperation.SetCell(
                                "Ops",
                                "A1",
                                new dev.erst.gridgrind.protocol.dto.CellInput.Text("x"))),
                        List.of())));
    assertEquals(
        GridGrindProblemCode.INVALID_REQUEST, unsupportedStreamingOperation.problem().code());
    assertTrue(unsupportedStreamingOperation.problem().message().contains("SET_CELL"));

    GridGrindResponse.Failure missingStreamingSheetMaterialization =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    new GridGrindRequest(
                        new GridGrindRequest.WorkbookSource.New(),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        new ExecutionModeInput(
                            ExecutionModeInput.ReadMode.FULL_XSSF,
                            ExecutionModeInput.WriteMode.STREAMING_WRITE),
                        null,
                        List.of(new WorkbookOperation.ForceFormulaRecalculationOnOpen()),
                        List.of())));
    assertEquals(
        GridGrindProblemCode.INVALID_REQUEST,
        missingStreamingSheetMaterialization.problem().code());
    assertTrue(
        missingStreamingSheetMaterialization
            .problem()
            .message()
            .contains("requires at least one ENSURE_SHEET or APPEND_ROW"));
  }

  private static GridGrindResponse.Success success(GridGrindResponse response) {
    return assertInstanceOf(GridGrindResponse.Success.class, response);
  }

  private static GridGrindResponse.Failure failure(GridGrindResponse response) {
    return assertInstanceOf(GridGrindResponse.Failure.class, response);
  }

  private static String savedPath(GridGrindResponse.Success success) {
    return switch (success.persistence()) {
      case GridGrindResponse.PersistenceOutcome.SavedAs savedAs -> savedAs.executionPath();
      case GridGrindResponse.PersistenceOutcome.Overwritten overwritten ->
          overwritten.executionPath();
      case GridGrindResponse.PersistenceOutcome.NotSaved _ ->
          throw new AssertionError("expected persisted workbook");
    };
  }
}
