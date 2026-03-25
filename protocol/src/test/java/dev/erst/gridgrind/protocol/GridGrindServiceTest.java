package dev.erst.gridgrind.protocol;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.ExcelCellValue;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.InvalidCellAddressException;
import dev.erst.gridgrind.excel.InvalidFormulaException;
import dev.erst.gridgrind.excel.SheetNotFoundException;
import dev.erst.gridgrind.excel.UnsupportedFormulaException;
import dev.erst.gridgrind.excel.WorkbookCommandExecutor;
import dev.erst.gridgrind.excel.WorkbookNotFoundException;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Integration tests for GridGrindService end-to-end request execution. */
class GridGrindServiceTest {
  @Test
  void executesAgentWorkflowAndReturnsWorkbookAnalysis() throws IOException {
    Path workbookPath = Files.createTempFile("gridgrind-agent-", ".xlsx");
    Files.deleteIfExists(workbookPath);

    GridGrindRequest request =
        new GridGrindRequest(
            new GridGrindRequest.WorkbookSource.New(),
            new GridGrindRequest.WorkbookPersistence.SaveAs(workbookPath.toString()),
            List.of(
                new WorkbookOperation.EnsureSheet("Budget"),
                new WorkbookOperation.AppendRow(
                    "Budget",
                    List.of(
                        new CellInput.Text("Item"),
                        new CellInput.Text("Amount"),
                        new CellInput.Text("Billable"))),
                new WorkbookOperation.AppendRow(
                    "Budget",
                    List.of(
                        new CellInput.Text("Hosting"),
                        new CellInput.Numeric(49.0),
                        new CellInput.BooleanValue(true))),
                new WorkbookOperation.AppendRow(
                    "Budget",
                    List.of(
                        new CellInput.Text("Domain"),
                        new CellInput.Numeric(12.0),
                        new CellInput.BooleanValue(false))),
                new WorkbookOperation.SetCell("Budget", "A4", new CellInput.Text("Total")),
                new WorkbookOperation.SetCell("Budget", "B4", new CellInput.Formula("SUM(B2:B3)")),
                new WorkbookOperation.AutoSizeColumns("Budget", List.of("A", "B", "C")),
                new WorkbookOperation.EvaluateFormulas(),
                new WorkbookOperation.ForceFormulaRecalculationOnOpen()),
            new GridGrindRequest.WorkbookAnalysisRequest(
                List.of(
                    new GridGrindRequest.SheetInspectionRequest(
                        "Budget", List.of("A1", "B4", "C2"), 4, 3))));

    GridGrindResponse response = new GridGrindService().execute(request);

    assertInstanceOf(GridGrindResponse.Success.class, response);
    GridGrindResponse.Success success = (GridGrindResponse.Success) response;
    assertEquals(GridGrindProtocolVersion.V1, success.protocolVersion());
    assertEquals(workbookPath.toAbsolutePath().toString(), success.savedWorkbookPath());
    assertTrue(Files.exists(workbookPath));
    assertEquals(1, success.workbook().sheetCount());
    assertEquals(List.of("Budget"), success.workbook().sheetNames());
    assertTrue(success.workbook().forceFormulaRecalculationOnOpen());
    assertEquals(1, success.sheets().size());
    assertEquals("Budget", success.sheets().get(0).sheetName());
    assertEquals("Item", success.sheets().get(0).requestedCells().get(0).stringValue());
    GridGrindResponse.CellReport.FormulaReport formulaCell =
        (GridGrindResponse.CellReport.FormulaReport)
            success.sheets().get(0).requestedCells().get(1);
    assertEquals("SUM(B2:B3)", formulaCell.formula());
    assertEquals(
        61.0, ((GridGrindResponse.CellReport.NumberReport) formulaCell.evaluation()).numberValue());
    assertTrue(success.sheets().get(0).requestedCells().get(2).booleanValue());
    assertEquals(4, success.sheets().get(0).previewRows().size());
  }

  @Test
  void opensExistingWorkbookAndOverwritesSource() throws IOException {
    Path workbookPath = Files.createTempFile("gridgrind-existing-", ".xlsx");

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Budget").setCell("A1", ExcelCellValue.text("Before"));
      workbook.save(workbookPath);
    }

    GridGrindRequest request =
        new GridGrindRequest(
            new GridGrindRequest.WorkbookSource.ExistingFile(workbookPath.toString()),
            new GridGrindRequest.WorkbookPersistence.OverwriteSource(),
            List.of(
                new WorkbookOperation.SetCell("Budget", "A1", new CellInput.Text("After")),
                new WorkbookOperation.SetCell("Budget", "B1", new CellInput.Numeric(12.0))),
            new GridGrindRequest.WorkbookAnalysisRequest(
                List.of(
                    new GridGrindRequest.SheetInspectionRequest(
                        "Budget", List.of("A1", "B1"), null, null))));

    GridGrindResponse response = new GridGrindService().execute(request);

    assertInstanceOf(GridGrindResponse.Success.class, response);
    GridGrindResponse.Success success = (GridGrindResponse.Success) response;
    assertEquals(workbookPath.toAbsolutePath().toString(), success.savedWorkbookPath());
    assertEquals("After", success.sheets().get(0).requestedCells().get(0).stringValue());
    assertEquals(12.0, success.sheets().get(0).requestedCells().get(1).numberValue());
    assertEquals(List.of(), success.sheets().get(0).previewRows());
  }

  @Test
  void returnsStructuredFailureForMissingWorkbookSource() {
    Path workbookPath = Path.of("missing-workbook.xlsx");

    GridGrindResponse response =
        new GridGrindService()
            .execute(
                new GridGrindRequest(
                    new GridGrindRequest.WorkbookSource.ExistingFile(workbookPath.toString()),
                    new GridGrindRequest.WorkbookPersistence.None(),
                    List.of(),
                    new GridGrindRequest.WorkbookAnalysisRequest(List.of())));

    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.WORKBOOK_NOT_FOUND, failure.problem().code());
    assertEquals("OPEN_WORKBOOK", failure.problem().context().stage());
    assertEquals(GridGrindProblemCategory.RESOURCE, failure.problem().category());
    assertEquals(GridGrindProblemRecovery.CHANGE_REQUEST, failure.problem().recovery());
    assertEquals(
        workbookPath.toAbsolutePath().toString(), failure.problem().context().sourceWorkbookPath());
  }

  @Test
  void returnsStructuredFailureWithOperationContext() {
    GridGrindResponse response =
        new GridGrindService()
            .execute(
                new GridGrindRequest(
                    new GridGrindRequest.WorkbookSource.New(),
                    new GridGrindRequest.WorkbookPersistence.None(),
                    List.of(new WorkbookOperation.AutoSizeColumns("Budget", List.of("A"))),
                    new GridGrindRequest.WorkbookAnalysisRequest(List.of())));

    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.SHEET_NOT_FOUND, failure.problem().code());
    assertEquals("APPLY_OPERATION", failure.problem().context().stage());
    assertEquals(0, failure.problem().context().operationIndex());
    assertEquals("AUTO_SIZE_COLUMNS", failure.problem().context().operationType());
    assertEquals("Budget", failure.problem().context().sheetName());
  }

  @Test
  void returnsStructuredFailureWithAnalysisContext() {
    GridGrindResponse response =
        new GridGrindService()
            .execute(
                new GridGrindRequest(
                    new GridGrindRequest.WorkbookSource.New(),
                    new GridGrindRequest.WorkbookPersistence.None(),
                    List.of(new WorkbookOperation.EnsureSheet("Budget")),
                    new GridGrindRequest.WorkbookAnalysisRequest(
                        List.of(
                            new GridGrindRequest.SheetInspectionRequest(
                                "Budget", List.of("A1"), null, null)))));

    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.CELL_NOT_FOUND, failure.problem().code());
    assertEquals("ANALYZE_WORKBOOK", failure.problem().context().stage());
    assertEquals("Budget", failure.problem().context().sheetName());
    assertEquals("A1", failure.problem().context().address());
  }

  @Test
  void returnsStructuredFailureWhenAnalysisTargetsMissingSheet() {
    GridGrindResponse response =
        new GridGrindService()
            .execute(
                new GridGrindRequest(
                    new GridGrindRequest.WorkbookSource.New(),
                    new GridGrindRequest.WorkbookPersistence.None(),
                    List.of(),
                    new GridGrindRequest.WorkbookAnalysisRequest(
                        List.of(
                            new GridGrindRequest.SheetInspectionRequest(
                                "Missing", List.of(), null, null)))));

    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.SHEET_NOT_FOUND, failure.problem().code());
    assertEquals("ANALYZE_WORKBOOK", failure.problem().context().stage());
    assertEquals("Missing", failure.problem().context().sheetName());
  }

  @Test
  void returnsStructuredFailureForInvalidPersistTarget() throws IOException {
    Path directory = Files.createTempDirectory("gridgrind-persist-target-");

    GridGrindResponse response =
        new GridGrindService()
            .execute(
                new GridGrindRequest(
                    new GridGrindRequest.WorkbookSource.New(),
                    new GridGrindRequest.WorkbookPersistence.SaveAs(directory.toString()),
                    List.of(new WorkbookOperation.EnsureSheet("Budget")),
                    new GridGrindRequest.WorkbookAnalysisRequest(List.of())));

    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.IO_ERROR, failure.problem().code());
    assertEquals("PERSIST_WORKBOOK", failure.problem().context().stage());
    assertEquals(
        directory.toAbsolutePath().toString(), failure.problem().context().persistencePath());
  }

  @Test
  void doesNotPersistWorkbookWhenAnalysisFails() throws IOException {
    Path workbookPath = Files.createTempFile("gridgrind-analysis-failure-", ".xlsx");
    Files.deleteIfExists(workbookPath);

    GridGrindResponse response =
        new GridGrindService()
            .execute(
                new GridGrindRequest(
                    new GridGrindRequest.WorkbookSource.New(),
                    new GridGrindRequest.WorkbookPersistence.SaveAs(workbookPath.toString()),
                    List.of(new WorkbookOperation.EnsureSheet("Budget")),
                    new GridGrindRequest.WorkbookAnalysisRequest(
                        List.of(
                            new GridGrindRequest.SheetInspectionRequest(
                                "Budget", List.of("A1"), null, null)))));

    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.CELL_NOT_FOUND, failure.problem().code());
    assertFalse(Files.exists(workbookPath));
  }

  @Test
  void returnsStructuredFailureForInvalidOverwriteSourceUsage() {
    GridGrindResponse response =
        new GridGrindService()
            .execute(
                new GridGrindRequest(
                    new GridGrindRequest.WorkbookSource.New(),
                    new GridGrindRequest.WorkbookPersistence.OverwriteSource(),
                    List.of(),
                    new GridGrindRequest.WorkbookAnalysisRequest(List.of())));

    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertEquals("PERSIST_WORKBOOK", failure.problem().context().stage());
  }

  @Test
  void returnsFormulaErrorForInvalidFormulaOperations() {
    GridGrindResponse response =
        new GridGrindService()
            .execute(
                new GridGrindRequest(
                    new GridGrindRequest.WorkbookSource.New(),
                    new GridGrindRequest.WorkbookPersistence.None(),
                    List.of(
                        new WorkbookOperation.SetCell("Data", "A1", new CellInput.Formula("SUM(")),
                        new WorkbookOperation.EvaluateFormulas()),
                    new GridGrindRequest.WorkbookAnalysisRequest(List.of())));

    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.INVALID_FORMULA, failure.problem().code());
    assertEquals(GridGrindProblemCategory.FORMULA, failure.problem().category());
    assertEquals("APPLY_OPERATION", failure.problem().context().stage());
    assertEquals("Invalid formula at Data!A1: SUM(", failure.problem().message());
    assertEquals("A1", failure.problem().context().address());
    assertEquals("SUM(", failure.problem().context().formula());
  }

  @Test
  void surfacesWorkbookFormulaLocationWhenEvaluationFails() {
    GridGrindResponse response =
        new GridGrindService()
            .execute(
                new GridGrindRequest(
                    new GridGrindRequest.WorkbookSource.New(),
                    new GridGrindRequest.WorkbookPersistence.None(),
                    List.of(
                        new WorkbookOperation.SetCell("Data", "A1", new CellInput.Numeric(1.0)),
                        new WorkbookOperation.SetCell("Data", "B1", new CellInput.Numeric(2.0)),
                        new WorkbookOperation.SetCell(
                            "Data", "C1", new CellInput.Formula("TEXTAFTER(\"a,b\",\",\")")),
                        new WorkbookOperation.EvaluateFormulas()),
                    new GridGrindRequest.WorkbookAnalysisRequest(List.of())));

    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.UNSUPPORTED_FORMULA, failure.problem().code());
    assertEquals("APPLY_OPERATION", failure.problem().context().stage());
    assertEquals(
        "Unsupported formula function TEXTAFTER at Data!C1: TEXTAFTER(\"a,b\",\",\")",
        failure.problem().message());
    assertEquals("Data", failure.problem().context().sheetName());
    assertEquals("C1", failure.problem().context().address());
    assertEquals("TEXTAFTER(\"a,b\",\",\")", failure.problem().context().formula());
  }

  @Test
  void returnsStructuredFailureWhenWorkbookCloseFailsAfterSuccess() {
    GridGrindService service =
        new GridGrindService(
            new WorkbookCommandExecutor(),
            workbook -> {
              throw new IOException("close failed");
            });

    GridGrindResponse response =
        service.execute(
            new GridGrindRequest(
                new GridGrindRequest.WorkbookSource.New(),
                new GridGrindRequest.WorkbookPersistence.None(),
                List.of(new WorkbookOperation.EnsureSheet("Budget")),
                new GridGrindRequest.WorkbookAnalysisRequest(List.of())));

    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.IO_ERROR, failure.problem().code());
    assertEquals("EXECUTE_REQUEST", failure.problem().context().stage());
    assertEquals("close failed", failure.problem().message());
    assertEquals(GridGrindProblemRecovery.CHECK_ENVIRONMENT, failure.problem().recovery());
  }

  @Test
  void preservesPrimaryFailureWhenWorkbookCloseAlsoFails() {
    GridGrindService service =
        new GridGrindService(
            new WorkbookCommandExecutor(),
            workbook -> {
              throw new IOException("close failed");
            });

    GridGrindResponse response =
        service.execute(
            new GridGrindRequest(
                new GridGrindRequest.WorkbookSource.New(),
                new GridGrindRequest.WorkbookPersistence.None(),
                List.of(new WorkbookOperation.AutoSizeColumns("Budget", List.of("A"))),
                new GridGrindRequest.WorkbookAnalysisRequest(List.of())));

    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.SHEET_NOT_FOUND, failure.problem().code());
    assertEquals("APPLY_OPERATION", failure.problem().context().stage());
    assertEquals(2, failure.problem().causes().size());
    assertTrue(
        failure
            .problem()
            .causes()
            .get(1)
            .message()
            .contains("Workbook close failed after the primary problem"));
    assertEquals("EXECUTE_REQUEST", failure.problem().causes().get(1).stage());
  }

  @Test
  void returnsStructuredFailureForNullRequests() {
    GridGrindResponse response = new GridGrindService().execute(null);

    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertEquals("VALIDATE_REQUEST", failure.problem().context().stage());
  }

  @Test
  void classifiesProblemCodesAndMessagesDeterministically() {
    assertEquals(
        GridGrindProblemCode.INVALID_CELL_ADDRESS,
        GridGrindService.problemCodeFor(
            new InvalidCellAddressException("BAD!", new IllegalArgumentException("bad"))));
    assertEquals(
        GridGrindProblemCode.INVALID_RANGE_ADDRESS,
        GridGrindService.problemCodeFor(
            new dev.erst.gridgrind.excel.InvalidRangeAddressException(
                "A1:", new IllegalArgumentException("bad"))));
    assertEquals(
        GridGrindProblemCode.INVALID_FORMULA,
        GridGrindService.problemCodeFor(
            new InvalidFormulaException(
                "Budget", "B4", "SUM(", "bad formula", new IllegalArgumentException("bad"))));
    assertEquals(
        GridGrindProblemCode.UNSUPPORTED_FORMULA,
        GridGrindService.problemCodeFor(
            new UnsupportedFormulaException(
                "Budget",
                "C1",
                "LAMBDA(x,x+1)(2)",
                "unsupported",
                new IllegalArgumentException("bad"))));
    assertEquals(
        GridGrindProblemCode.WORKBOOK_NOT_FOUND,
        GridGrindService.problemCodeFor(new WorkbookNotFoundException(Path.of("missing.xlsx"))));
    assertEquals(
        GridGrindProblemCode.SHEET_NOT_FOUND,
        GridGrindService.problemCodeFor(new SheetNotFoundException("Budget")));
    assertEquals(
        GridGrindProblemCode.IO_ERROR, GridGrindService.problemCodeFor(new IOException("disk")));
    assertEquals(
        GridGrindProblemCode.INVALID_REQUEST,
        GridGrindService.problemCodeFor(new IllegalArgumentException("bad request")));
    assertEquals(
        GridGrindProblemCode.INTERNAL_ERROR,
        GridGrindService.problemCodeFor(new UnsupportedOperationException()));
    assertEquals(
        "bad request", GridGrindService.messageFor(new IllegalArgumentException("bad request")));
    assertEquals(
        "UnsupportedOperationException",
        GridGrindService.messageFor(new UnsupportedOperationException()));
  }

  @Test
  void executesRangeAndStyleWorkflowAndSurfacesStyledCells() throws IOException {
    Path workbookPath = Files.createTempFile("gridgrind-range-style-", ".xlsx");
    Files.deleteIfExists(workbookPath);

    GridGrindResponse response =
        new GridGrindService()
            .execute(
                new GridGrindRequest(
                    new GridGrindRequest.WorkbookSource.New(),
                    new GridGrindRequest.WorkbookPersistence.SaveAs(workbookPath.toString()),
                    List.of(
                        new WorkbookOperation.EnsureSheet("Budget"),
                        new WorkbookOperation.SetRange(
                            "Budget",
                            "A1:B2",
                            List.of(
                                List.of(new CellInput.Text("Item"), new CellInput.Text("Amount")),
                                List.of(
                                    new CellInput.Text("Hosting"), new CellInput.Numeric(49.0)))),
                        new WorkbookOperation.ApplyStyle(
                            "Budget",
                            "A1:B1",
                            new CellStyleInput(
                                "#,##0.00",
                                true,
                                null,
                                true,
                                CellStyleInput.HorizontalAlignmentInput.CENTER,
                                CellStyleInput.VerticalAlignmentInput.CENTER)),
                        new WorkbookOperation.ApplyStyle(
                            "Budget",
                            "C1",
                            new CellStyleInput(
                                null,
                                null,
                                true,
                                null,
                                CellStyleInput.HorizontalAlignmentInput.RIGHT,
                                CellStyleInput.VerticalAlignmentInput.BOTTOM)),
                        new WorkbookOperation.SetCell(
                            "Budget", "B3", new CellInput.Formula("SUM(B2:B2)")),
                        new WorkbookOperation.EvaluateFormulas(),
                        new WorkbookOperation.ClearRange("Budget", "A2")),
                    new GridGrindRequest.WorkbookAnalysisRequest(
                        List.of(
                            new GridGrindRequest.SheetInspectionRequest(
                                "Budget", List.of("A1", "A2", "B3", "C1"), 3, 3)))));

    assertInstanceOf(GridGrindResponse.Success.class, response);
    GridGrindResponse.Success success = (GridGrindResponse.Success) response;
    assertTrue(Files.exists(workbookPath));
    assertEquals(
        "CENTER",
        success.sheets().getFirst().requestedCells().getFirst().style().horizontalAlignment());
    assertTrue(success.sheets().getFirst().requestedCells().getFirst().style().bold());
    assertTrue(success.sheets().getFirst().requestedCells().getFirst().style().wrapText());
    assertEquals("BLANK", success.sheets().getFirst().requestedCells().get(1).effectiveType());
    assertEquals("49", success.sheets().getFirst().requestedCells().get(2).displayValue());
    assertTrue(
        success.sheets().getFirst().previewRows().getFirst().cells().stream()
            .anyMatch(cell -> "C1".equals(cell.address())));
    assertTrue(success.sheets().getFirst().requestedCells().get(3).style().italic());
  }

  @Test
  void producesErrorReportForCellsWithErrorValues() {
    GridGrindResponse response =
        new GridGrindService()
            .execute(
                new GridGrindRequest(
                    new GridGrindRequest.WorkbookSource.New(),
                    new GridGrindRequest.WorkbookPersistence.None(),
                    List.of(
                        new WorkbookOperation.EnsureSheet("Data"),
                        new WorkbookOperation.SetCell("Data", "A1", new CellInput.Formula("1/0")),
                        new WorkbookOperation.EvaluateFormulas()),
                    new GridGrindRequest.WorkbookAnalysisRequest(
                        List.of(
                            new GridGrindRequest.SheetInspectionRequest(
                                "Data", List.of("A1"), null, null)))));

    assertInstanceOf(GridGrindResponse.Success.class, response);
    GridGrindResponse.Success success = (GridGrindResponse.Success) response;
    GridGrindResponse.CellReport cell = success.sheets().get(0).requestedCells().get(0);
    assertInstanceOf(GridGrindResponse.CellReport.FormulaReport.class, cell);
    GridGrindResponse.CellReport evaluation =
        ((GridGrindResponse.CellReport.FormulaReport) cell).evaluation();
    assertInstanceOf(GridGrindResponse.CellReport.ErrorReport.class, evaluation);
    assertEquals("ERROR", evaluation.effectiveType());
  }

  @Test
  void extractsFormulaFromSetCellOperationWhenExceptionCarriesNone() {
    // Uses an invalid cell address so SetCell fails with InvalidCellAddressException,
    // which carries no formula. The service must extract the formula from the operation itself.
    GridGrindResponse response =
        new GridGrindService()
            .execute(
                new GridGrindRequest(
                    new GridGrindRequest.WorkbookSource.New(),
                    new GridGrindRequest.WorkbookPersistence.None(),
                    List.of(
                        new WorkbookOperation.EnsureSheet("Data"),
                        new WorkbookOperation.SetCell(
                            "Data", "INVALID!", new CellInput.Formula("SUM(B1:B2)"))),
                    new GridGrindRequest.WorkbookAnalysisRequest(List.of())));

    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals("APPLY_OPERATION", failure.problem().context().stage());
    // The exception carries no formula; the service extracts it from the SetCell operation itself.
    assertEquals("SUM(B1:B2)", failure.problem().context().formula());
  }

  @Test
  void persistencePathResolvesCorrectlyForAllPersistenceAndSourceCombinations() {
    GridGrindService service = new GridGrindService();

    GridGrindRequest.WorkbookSource newSource = new GridGrindRequest.WorkbookSource.New();
    GridGrindRequest.WorkbookSource existingFile =
        new GridGrindRequest.WorkbookSource.ExistingFile("/tmp/source.xlsx");
    GridGrindRequest.WorkbookPersistence none = new GridGrindRequest.WorkbookPersistence.None();
    GridGrindRequest.WorkbookPersistence overwrite =
        new GridGrindRequest.WorkbookPersistence.OverwriteSource();
    GridGrindRequest.WorkbookPersistence saveAs =
        new GridGrindRequest.WorkbookPersistence.SaveAs("/tmp/out.xlsx");

    // SaveAs always resolves to the saveAs path regardless of source
    assertEquals(
        Path.of("/tmp/out.xlsx").toAbsolutePath().toString(),
        service.persistencePath(newSource, saveAs));

    // OverwriteSource + ExistingFile resolves to the source path
    assertEquals(
        Path.of("/tmp/source.xlsx").toAbsolutePath().toString(),
        service.persistencePath(existingFile, overwrite));

    // OverwriteSource + New resolves to null (no source path to overwrite)
    assertNull(service.persistencePath(newSource, overwrite));

    // None always resolves to null regardless of source
    assertNull(service.persistencePath(newSource, none));
    // Note: None + ExistingFile also resolves to null since None persistence has no target path.
    // This arm is only exercised defensively (persistWorkbook with None cannot fail in practice).
    assertNull(service.persistencePath(existingFile, none));
  }

  @Test
  void guardsCatastrophicRuntimeExceptionsAndProducesExecuteRequestFailure() {
    // The closer throws RuntimeException on the first call (inside the success path lambda),
    // which bubbles out of closeWorkbook (which only catches IOException), through the lambda,
    // and is caught by guardUnexpectedRuntime. The second closeWorkbook call in the catch
    // block succeeds so the failure response is returned cleanly.
    int[] callCount = {0};
    GridGrindService service =
        new GridGrindService(
            new WorkbookCommandExecutor(),
            workbook -> {
              int count = callCount[0];
              callCount[0] = count + 1;
              if (count == 0) {
                throw new IllegalStateException("catastrophic close failure");
              }
            });

    GridGrindResponse response =
        service.execute(
            new GridGrindRequest(
                new GridGrindRequest.WorkbookSource.New(),
                new GridGrindRequest.WorkbookPersistence.None(),
                List.of(new WorkbookOperation.EnsureSheet("Budget")),
                new GridGrindRequest.WorkbookAnalysisRequest(List.of())));

    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.INTERNAL_ERROR, failure.problem().code());
    assertEquals("EXECUTE_REQUEST", failure.problem().context().stage());
  }

  @Test
  void returnsStructuredFailureForInvalidRangeOperations() {
    GridGrindResponse response =
        new GridGrindService()
            .execute(
                new GridGrindRequest(
                    new GridGrindRequest.WorkbookSource.New(),
                    new GridGrindRequest.WorkbookPersistence.None(),
                    List.of(
                        new WorkbookOperation.EnsureSheet("Budget"),
                        new WorkbookOperation.ClearRange("Budget", "A1:")),
                    new GridGrindRequest.WorkbookAnalysisRequest(List.of())));

    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.INVALID_RANGE_ADDRESS, failure.problem().code());
    assertEquals("APPLY_OPERATION", failure.problem().context().stage());
    assertEquals("Budget", failure.problem().context().sheetName());
    assertEquals("A1:", failure.problem().context().range());
    assertEquals("CLEAR_RANGE", failure.problem().context().operationType());
  }
}
