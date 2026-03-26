package dev.erst.gridgrind.protocol;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.ExcelBorderStyle;
import dev.erst.gridgrind.excel.ExcelCellValue;
import dev.erst.gridgrind.excel.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.ExcelVerticalAlignment;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.InvalidCellAddressException;
import dev.erst.gridgrind.excel.InvalidFormulaException;
import dev.erst.gridgrind.excel.SheetNotFoundException;
import dev.erst.gridgrind.excel.UnsupportedFormulaException;
import dev.erst.gridgrind.excel.WorkbookCommandExecutor;
import dev.erst.gridgrind.excel.WorkbookNotFoundException;
import dev.erst.gridgrind.excel.XlsxRoundTrip;
import java.io.IOException;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
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
                new WorkbookOperation.AutoSizeColumns("Budget"),
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
  void managesSheetsAndPersistsUpdatedSheetOrder() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-sheet-ops-");

    GridGrindRequest request =
        new GridGrindRequest(
            new GridGrindRequest.WorkbookSource.New(),
            new GridGrindRequest.WorkbookPersistence.SaveAs(workbookPath.toString()),
            List.of(
                new WorkbookOperation.EnsureSheet("Budget"),
                new WorkbookOperation.EnsureSheet("Archive"),
                new WorkbookOperation.EnsureSheet("Scratch"),
                new WorkbookOperation.SetCell("Archive", "A1", new CellInput.Text("Old")),
                new WorkbookOperation.SetCell("Budget", "A1", new CellInput.Text("Live")),
                new WorkbookOperation.RenameSheet("Archive", "History"),
                new WorkbookOperation.MoveSheet("History", 0),
                new WorkbookOperation.DeleteSheet("Scratch")),
            new GridGrindRequest.WorkbookAnalysisRequest(
                List.of(
                    new GridGrindRequest.SheetInspectionRequest(
                        "History", List.of("A1"), null, null),
                    new GridGrindRequest.SheetInspectionRequest(
                        "Budget", List.of("A1"), null, null))));

    GridGrindResponse response = new GridGrindService().execute(request);

    assertInstanceOf(GridGrindResponse.Success.class, response);
    GridGrindResponse.Success success = (GridGrindResponse.Success) response;
    assertEquals(workbookPath.toAbsolutePath().toString(), success.savedWorkbookPath());
    assertEquals(2, success.workbook().sheetCount());
    assertEquals(List.of("History", "Budget"), success.workbook().sheetNames());
    assertEquals("History", success.sheets().get(0).sheetName());
    assertEquals("Old", success.sheets().get(0).requestedCells().get(0).stringValue());
    assertEquals("Budget", success.sheets().get(1).sheetName());
    assertEquals("Live", success.sheets().get(1).requestedCells().get(0).stringValue());
    assertEquals(List.of("History", "Budget"), XlsxRoundTrip.sheetOrder(workbookPath));
  }

  @Test
  void appliesStructuralLayoutOperationsAndPersistsWorkbookShape() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-layout-ops-");

    GridGrindRequest request =
        new GridGrindRequest(
            new GridGrindRequest.WorkbookSource.New(),
            new GridGrindRequest.WorkbookPersistence.SaveAs(workbookPath.toString()),
            List.of(
                new WorkbookOperation.EnsureSheet("Budget"),
                new WorkbookOperation.SetCell("Budget", "A1", new CellInput.Text("Quarterly")),
                new WorkbookOperation.MergeCells("Budget", "A1:B1"),
                new WorkbookOperation.SetColumnWidth("Budget", 0, 1, 16.0),
                new WorkbookOperation.SetRowHeight("Budget", 0, 0, 28.5),
                new WorkbookOperation.FreezePanes("Budget", 1, 1, 1, 1)),
            new GridGrindRequest.WorkbookAnalysisRequest(
                List.of(
                    new GridGrindRequest.SheetInspectionRequest(
                        "Budget", List.of("A1"), null, null))));

    GridGrindResponse response = new GridGrindService().execute(request);

    assertInstanceOf(GridGrindResponse.Success.class, response);
    GridGrindResponse.Success success = (GridGrindResponse.Success) response;
    assertEquals(workbookPath.toAbsolutePath().toString(), success.savedWorkbookPath());
    assertEquals(
        "Quarterly", success.sheets().getFirst().requestedCells().getFirst().stringValue());
    assertEquals(List.of("A1:B1"), XlsxRoundTrip.mergedRegions(workbookPath, "Budget"));
    assertEquals(4096, XlsxRoundTrip.columnWidth(workbookPath, "Budget", 0));
    assertEquals((short) 570, XlsxRoundTrip.rowHeightTwips(workbookPath, "Budget", 0));
    assertEquals(
        new XlsxRoundTrip.FreezePaneState.Frozen(1, 1, 1, 1),
        XlsxRoundTrip.freezePaneState(workbookPath, "Budget"));
  }

  @Test
  void returnsStructuredFailureWhenMoveSheetTargetsMissingSheet() {
    GridGrindResponse response =
        new GridGrindService()
            .execute(
                new GridGrindRequest(
                    new GridGrindRequest.WorkbookSource.New(),
                    new GridGrindRequest.WorkbookPersistence.None(),
                    List.of(new WorkbookOperation.MoveSheet("Missing", 0)),
                    new GridGrindRequest.WorkbookAnalysisRequest(List.of())));

    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.SHEET_NOT_FOUND, failure.problem().code());
    assertEquals("APPLY_OPERATION", failure.problem().context().stage());
    assertEquals("MOVE_SHEET", failure.problem().context().operationType());
    assertEquals("Missing", failure.problem().context().sheetName());
  }

  @Test
  void returnsStructuredFailureForConflictingRenameTarget() {
    GridGrindResponse response =
        new GridGrindService()
            .execute(
                new GridGrindRequest(
                    new GridGrindRequest.WorkbookSource.New(),
                    new GridGrindRequest.WorkbookPersistence.None(),
                    List.of(
                        new WorkbookOperation.EnsureSheet("Budget"),
                        new WorkbookOperation.EnsureSheet("Summary"),
                        new WorkbookOperation.RenameSheet("Budget", "Summary")),
                    new GridGrindRequest.WorkbookAnalysisRequest(List.of())));

    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertEquals("APPLY_OPERATION", failure.problem().context().stage());
    assertEquals(2, failure.problem().context().operationIndex());
    assertEquals("RENAME_SHEET", failure.problem().context().operationType());
    assertEquals("Budget", failure.problem().context().sheetName());
    assertEquals("Sheet already exists: Summary", failure.problem().message());
  }

  @Test
  void returnsStructuredFailureForInvalidMoveSheetIndex() {
    GridGrindResponse response =
        new GridGrindService()
            .execute(
                new GridGrindRequest(
                    new GridGrindRequest.WorkbookSource.New(),
                    new GridGrindRequest.WorkbookPersistence.None(),
                    List.of(
                        new WorkbookOperation.EnsureSheet("Budget"),
                        new WorkbookOperation.MoveSheet("Budget", 1)),
                    new GridGrindRequest.WorkbookAnalysisRequest(List.of())));

    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertEquals("APPLY_OPERATION", failure.problem().context().stage());
    assertEquals(1, failure.problem().context().operationIndex());
    assertEquals("MOVE_SHEET", failure.problem().context().operationType());
    assertEquals("Budget", failure.problem().context().sheetName());
    assertEquals("targetIndex must be between 0 and 0 (inclusive): 1", failure.problem().message());
  }

  @Test
  void returnsStructuredFailureForMergeCellsWithInvalidRange() {
    GridGrindResponse response =
        new GridGrindService()
            .execute(
                new GridGrindRequest(
                    new GridGrindRequest.WorkbookSource.New(),
                    new GridGrindRequest.WorkbookPersistence.None(),
                    List.of(
                        new WorkbookOperation.EnsureSheet("Budget"),
                        new WorkbookOperation.MergeCells("Budget", "A1:")),
                    new GridGrindRequest.WorkbookAnalysisRequest(List.of())));

    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.INVALID_RANGE_ADDRESS, failure.problem().code());
    assertEquals("APPLY_OPERATION", failure.problem().context().stage());
    assertEquals("MERGE_CELLS", failure.problem().context().operationType());
    assertEquals("Budget", failure.problem().context().sheetName());
    assertEquals("A1:", failure.problem().context().range());
  }

  @Test
  void returnsStructuredFailureForUnmergeCellsWithoutExactMatch() {
    GridGrindResponse response =
        new GridGrindService()
            .execute(
                new GridGrindRequest(
                    new GridGrindRequest.WorkbookSource.New(),
                    new GridGrindRequest.WorkbookPersistence.None(),
                    List.of(
                        new WorkbookOperation.EnsureSheet("Budget"),
                        new WorkbookOperation.UnmergeCells("Budget", "A1:B2")),
                    new GridGrindRequest.WorkbookAnalysisRequest(List.of())));

    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertEquals("APPLY_OPERATION", failure.problem().context().stage());
    assertEquals("UNMERGE_CELLS", failure.problem().context().operationType());
    assertEquals("Budget", failure.problem().context().sheetName());
    assertEquals("A1:B2", failure.problem().context().range());
    assertEquals("No merged region matches range: A1:B2", failure.problem().message());
  }

  @Test
  void returnsStructuredFailureWhenFreezePanesTargetsMissingSheet() {
    GridGrindResponse response =
        new GridGrindService()
            .execute(
                new GridGrindRequest(
                    new GridGrindRequest.WorkbookSource.New(),
                    new GridGrindRequest.WorkbookPersistence.None(),
                    List.of(new WorkbookOperation.FreezePanes("Missing", 1, 1, 1, 1)),
                    new GridGrindRequest.WorkbookAnalysisRequest(List.of())));

    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.SHEET_NOT_FOUND, failure.problem().code());
    assertEquals("APPLY_OPERATION", failure.problem().context().stage());
    assertEquals("FREEZE_PANES", failure.problem().context().operationType());
    assertEquals("Missing", failure.problem().context().sheetName());
  }

  @Test
  void preservesStructuralWorkbookStateAcrossExistingWorkbookRoundTrips() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-round-trip-");

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      var budget = workbook.createSheet("Budget");
      workbook.createSheet("Summary");

      budget.addMergedRegion(CellRangeAddress.valueOf("A1:B2"));
      budget.setColumnWidth(0, 4096);
      budget.createRow(0).setHeightInPoints(28.5f);
      budget.createFreezePane(1, 2, 3, 4);
      budget.createRow(2).createCell(2).setCellValue("Before");

      try (var outputStream = Files.newOutputStream(workbookPath)) {
        workbook.write(outputStream);
      }
    }

    GridGrindRequest request =
        new GridGrindRequest(
            new GridGrindRequest.WorkbookSource.ExistingFile(workbookPath.toString()),
            new GridGrindRequest.WorkbookPersistence.OverwriteSource(),
            List.of(new WorkbookOperation.SetCell("Budget", "C3", new CellInput.Text("After"))),
            new GridGrindRequest.WorkbookAnalysisRequest(
                List.of(
                    new GridGrindRequest.SheetInspectionRequest(
                        "Budget", List.of("C3"), null, null))));

    GridGrindResponse response = new GridGrindService().execute(request);

    assertInstanceOf(GridGrindResponse.Success.class, response);
    GridGrindResponse.Success success = (GridGrindResponse.Success) response;
    assertEquals(workbookPath.toAbsolutePath().toString(), success.savedWorkbookPath());
    assertEquals("After", success.sheets().getFirst().requestedCells().getFirst().stringValue());
    assertEquals(List.of("Budget", "Summary"), XlsxRoundTrip.sheetOrder(workbookPath));
    assertEquals(List.of("A1:B2"), XlsxRoundTrip.mergedRegions(workbookPath, "Budget"));
    assertEquals(4096, XlsxRoundTrip.columnWidth(workbookPath, "Budget", 0));
    assertEquals((short) 570, XlsxRoundTrip.rowHeightTwips(workbookPath, "Budget", 0));
    assertEquals(
        new XlsxRoundTrip.FreezePaneState.Frozen(1, 2, 3, 4),
        XlsxRoundTrip.freezePaneState(workbookPath, "Budget"));
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
                    List.of(new WorkbookOperation.AutoSizeColumns("Budget")),
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
    Path parentFile = Files.createTempFile("gridgrind-persist-target-", ".tmp");
    Path workbookPath = parentFile.resolve("book.xlsx");

    GridGrindResponse response =
        new GridGrindService()
            .execute(
                new GridGrindRequest(
                    new GridGrindRequest.WorkbookSource.New(),
                    new GridGrindRequest.WorkbookPersistence.SaveAs(workbookPath.toString()),
                    List.of(new WorkbookOperation.EnsureSheet("Budget")),
                    new GridGrindRequest.WorkbookAnalysisRequest(List.of())));

    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.IO_ERROR, failure.problem().code());
    assertEquals("PERSIST_WORKBOOK", failure.problem().context().stage());
    assertEquals(
        workbookPath.toAbsolutePath().toString(), failure.problem().context().persistencePath());
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
    assertEquals("VALIDATE_REQUEST", failure.problem().context().stage());
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
  void returnsFormulaErrorForMalformedParserStateFormulaOperations() {
    GridGrindResponse response =
        new GridGrindService()
            .execute(
                new GridGrindRequest(
                    new GridGrindRequest.WorkbookSource.New(),
                    new GridGrindRequest.WorkbookPersistence.None(),
                    List.of(
                        new WorkbookOperation.SetCell(
                            "Data", "A1", new CellInput.Formula("[^owe_e`ffffff"))),
                    new GridGrindRequest.WorkbookAnalysisRequest(List.of())));

    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals(GridGrindProblemCode.INVALID_FORMULA, failure.problem().code());
    assertEquals(GridGrindProblemCategory.FORMULA, failure.problem().category());
    assertEquals("APPLY_OPERATION", failure.problem().context().stage());
    assertEquals("Invalid formula at Data!A1: [^owe_e`ffffff", failure.problem().message());
    assertEquals("A1", failure.problem().context().address());
    assertEquals("[^owe_e`ffffff", failure.problem().context().formula());
    assertEquals(
        "Invalid formula at Data!A1: [^owe_e`ffffff",
        failure.problem().causes().getFirst().message());
    assertEquals("InvalidFormulaException", failure.problem().causes().getFirst().type());
    assertEquals(
        "Parsed past the end of the formula, pos: 15, length: 14, formula: [^owe_e`ffffff",
        failure.problem().causes().get(1).message());
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
                List.of(new WorkbookOperation.AutoSizeColumns("Budget")),
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
                                ExcelHorizontalAlignment.CENTER,
                                ExcelVerticalAlignment.CENTER,
                                null,
                                null,
                                null,
                                null,
                                null,
                                null,
                                null)),
                        new WorkbookOperation.ApplyStyle(
                            "Budget",
                            "C1",
                            new CellStyleInput(
                                null,
                                null,
                                true,
                                null,
                                ExcelHorizontalAlignment.RIGHT,
                                ExcelVerticalAlignment.BOTTOM,
                                null,
                                null,
                                null,
                                null,
                                null,
                                null,
                                null)),
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
        ExcelHorizontalAlignment.CENTER,
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
  void executesFormattingDepthWorkflowAndPersistsReportedStyleState() throws IOException {
    Path workbookPath = Files.createTempFile("gridgrind-format-depth-", ".xlsx");
    Files.deleteIfExists(workbookPath);

    GridGrindResponse response =
        new GridGrindService()
            .execute(
                new GridGrindRequest(
                    new GridGrindRequest.WorkbookSource.New(),
                    new GridGrindRequest.WorkbookPersistence.SaveAs(workbookPath.toString()),
                    List.of(
                        new WorkbookOperation.EnsureSheet("Budget"),
                        new WorkbookOperation.SetCell("Budget", "A1", new CellInput.Text("Item")),
                        new WorkbookOperation.ApplyStyle(
                            "Budget",
                            "A1",
                            new CellStyleInput(
                                null,
                                true,
                                false,
                                true,
                                ExcelHorizontalAlignment.CENTER,
                                ExcelVerticalAlignment.TOP,
                                "Aptos",
                                new FontHeightInput.Points(new BigDecimal("11.5")),
                                "#1F4E78",
                                true,
                                true,
                                "#FFF2CC",
                                new CellBorderInput(
                                    new CellBorderSideInput(ExcelBorderStyle.THIN),
                                    null,
                                    new CellBorderSideInput(ExcelBorderStyle.DOUBLE),
                                    null,
                                    null)))),
                    new GridGrindRequest.WorkbookAnalysisRequest(
                        List.of(
                            new GridGrindRequest.SheetInspectionRequest(
                                "Budget", List.of("A1"), null, null)))));

    assertInstanceOf(GridGrindResponse.Success.class, response);
    GridGrindResponse.Success success = (GridGrindResponse.Success) response;
    GridGrindResponse.CellStyleReport style =
        success.sheets().getFirst().requestedCells().getFirst().style();
    assertTrue(Files.exists(workbookPath));
    assertTrue(style.bold());
    assertFalse(style.italic());
    assertTrue(style.wrapText());
    assertEquals(ExcelHorizontalAlignment.CENTER, style.horizontalAlignment());
    assertEquals(ExcelVerticalAlignment.TOP, style.verticalAlignment());
    assertEquals("Aptos", style.fontName());
    assertEquals(230, style.fontHeight().twips());
    assertEquals(new BigDecimal("11.5"), style.fontHeight().points());
    assertEquals("#1F4E78", style.fontColor());
    assertTrue(style.underline());
    assertTrue(style.strikeout());
    assertEquals("#FFF2CC", style.fillColor());
    assertEquals(ExcelBorderStyle.THIN, style.topBorderStyle());
    assertEquals(ExcelBorderStyle.DOUBLE, style.rightBorderStyle());
    assertEquals(ExcelBorderStyle.THIN, style.bottomBorderStyle());
    assertEquals(ExcelBorderStyle.THIN, style.leftBorderStyle());

    assertEquals(
        style, toResponseStyleReport(XlsxRoundTrip.cellStyle(workbookPath, "Budget", "A1")));
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

  private GridGrindResponse.CellStyleReport toResponseStyleReport(
      dev.erst.gridgrind.excel.ExcelCellStyleSnapshot style) {
    return new GridGrindResponse.CellStyleReport(
        style.numberFormat(),
        style.bold(),
        style.italic(),
        style.wrapText(),
        style.horizontalAlignment(),
        style.verticalAlignment(),
        style.fontName(),
        FontHeightReport.fromExcelFontHeight(style.fontHeight()),
        style.fontColor(),
        style.underline(),
        style.strikeout(),
        style.fillColor(),
        style.topBorderStyle(),
        style.rightBorderStyle(),
        style.bottomBorderStyle(),
        style.leftBorderStyle());
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

  @Test
  void returnsStructuredFailureForSetRangeWithInvalidRange() {
    GridGrindResponse response =
        new GridGrindService()
            .execute(
                new GridGrindRequest(
                    new GridGrindRequest.WorkbookSource.New(),
                    new GridGrindRequest.WorkbookPersistence.None(),
                    List.of(
                        new WorkbookOperation.EnsureSheet("Budget"),
                        new WorkbookOperation.SetRange(
                            "Budget", "A1:", List.of(List.of(new CellInput.Text("x"))))),
                    new GridGrindRequest.WorkbookAnalysisRequest(List.of())));

    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals("APPLY_OPERATION", failure.problem().context().stage());
    assertEquals("SET_RANGE", failure.problem().context().operationType());
    assertEquals("Budget", failure.problem().context().sheetName());
    assertEquals("A1:", failure.problem().context().range());
  }

  @Test
  void returnsStructuredFailureForApplyStyleWithInvalidRange() {
    GridGrindResponse response =
        new GridGrindService()
            .execute(
                new GridGrindRequest(
                    new GridGrindRequest.WorkbookSource.New(),
                    new GridGrindRequest.WorkbookPersistence.None(),
                    List.of(
                        new WorkbookOperation.EnsureSheet("Budget"),
                        new WorkbookOperation.ApplyStyle(
                            "Budget",
                            "A1:",
                            new CellStyleInput(
                                null, true, null, null, null, null, null, null, null, null, null,
                                null, null))),
                    new GridGrindRequest.WorkbookAnalysisRequest(List.of())));

    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals("APPLY_OPERATION", failure.problem().context().stage());
    assertEquals("APPLY_STYLE", failure.problem().context().operationType());
    assertEquals("Budget", failure.problem().context().sheetName());
    assertEquals("A1:", failure.problem().context().range());
  }

  @Test
  void returnsStructuredFailureForAppendRowWithInvalidFormula() {
    GridGrindResponse response =
        new GridGrindService()
            .execute(
                new GridGrindRequest(
                    new GridGrindRequest.WorkbookSource.New(),
                    new GridGrindRequest.WorkbookPersistence.None(),
                    List.of(
                        new WorkbookOperation.EnsureSheet("Budget"),
                        new WorkbookOperation.AppendRow(
                            "Budget", List.of(new CellInput.Formula("SUM(")))),
                    new GridGrindRequest.WorkbookAnalysisRequest(List.of())));

    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals("APPLY_OPERATION", failure.problem().context().stage());
    assertEquals("APPEND_ROW", failure.problem().context().operationType());
    assertEquals("Budget", failure.problem().context().sheetName());
  }

  @Test
  void returnsStructuredFailureForEnsureSheetWithInvalidSheetName() {
    // Sheet names exceeding Excel's 31-character limit are rejected by POI.
    GridGrindResponse response =
        new GridGrindService()
            .execute(
                new GridGrindRequest(
                    new GridGrindRequest.WorkbookSource.New(),
                    new GridGrindRequest.WorkbookPersistence.None(),
                    List.of(new WorkbookOperation.EnsureSheet("[Budget]")),
                    new GridGrindRequest.WorkbookAnalysisRequest(List.of())));

    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals("APPLY_OPERATION", failure.problem().context().stage());
    assertEquals("ENSURE_SHEET", failure.problem().context().operationType());
    assertEquals("[Budget]", failure.problem().context().sheetName());
  }

  @Test
  void extractsNullContextForOperationsWithNoSheetAddressRangeOrFormula() {
    // Direct unit tests for the package-private context-extraction helpers.
    // These operation types (ForceFormulaRecalculationOnOpen, EvaluateFormulas without
    // FormulaException, AppendRow without FormulaException) carry no sheet/address/range/formula
    // in the operation record and produce exceptions that do not carry formula context.
    RuntimeException ex = new RuntimeException("test");
    WorkbookOperation forceRecalc = new WorkbookOperation.ForceFormulaRecalculationOnOpen();
    WorkbookOperation evalFormulas = new WorkbookOperation.EvaluateFormulas();
    WorkbookOperation appendRow =
        new WorkbookOperation.AppendRow("Budget", List.of(new CellInput.Text("x")));

    WorkbookOperation ensureSheet = new WorkbookOperation.EnsureSheet("Budget");

    assertNull(GridGrindService.formulaFor(forceRecalc, ex));
    assertNull(GridGrindService.formulaFor(evalFormulas, ex));
    assertNull(GridGrindService.formulaFor(appendRow, ex));
    assertNull(GridGrindService.formulaFor(ensureSheet, ex));

    assertNull(GridGrindService.sheetNameFor(forceRecalc, ex));
    assertNull(GridGrindService.sheetNameFor(evalFormulas, ex));
    assertNull(GridGrindService.addressFor(forceRecalc, ex));
    assertNull(GridGrindService.addressFor(evalFormulas, ex));
    assertNull(GridGrindService.rangeFor(forceRecalc, ex));
    assertNull(GridGrindService.rangeFor(evalFormulas, ex));
  }

  @Test
  void extractsNullFormulaFromSetCellWithNonFormulaValueWhenExceptionCarriesNone() {
    // SetCell fails with an invalid address; value is Text (not Formula). The formula context
    // must be null since neither the exception nor the operation carries a formula.
    GridGrindResponse response =
        new GridGrindService()
            .execute(
                new GridGrindRequest(
                    new GridGrindRequest.WorkbookSource.New(),
                    new GridGrindRequest.WorkbookPersistence.None(),
                    List.of(
                        new WorkbookOperation.EnsureSheet("Budget"),
                        new WorkbookOperation.SetCell(
                            "Budget", "INVALID!", new CellInput.Text("hello"))),
                    new GridGrindRequest.WorkbookAnalysisRequest(List.of())));

    assertInstanceOf(GridGrindResponse.Failure.class, response);
    GridGrindResponse.Failure failure = (GridGrindResponse.Failure) response;
    assertEquals("APPLY_OPERATION", failure.problem().context().stage());
    assertEquals("SET_CELL", failure.problem().context().operationType());
    assertNull(failure.problem().context().formula());
  }

  @Test
  void formulaForSetCellReturnsNullForAllNonFormulaValueTypes() {
    // Direct unit tests covering each CellInput subtype in the SetCell arm of formulaFor.
    // Text is covered by extractsNullFormulaFromSetCellWithNonFormulaValueWhenExceptionCarriesNone;
    // this test covers the remaining five non-formula types.
    RuntimeException ex = new RuntimeException("test");
    assertNull(
        GridGrindService.formulaFor(
            new WorkbookOperation.SetCell("S", "A1", new CellInput.Blank()), ex));
    assertNull(
        GridGrindService.formulaFor(
            new WorkbookOperation.SetCell("S", "A1", new CellInput.Numeric(1.0)), ex));
    assertNull(
        GridGrindService.formulaFor(
            new WorkbookOperation.SetCell("S", "A1", new CellInput.BooleanValue(true)), ex));
    assertNull(
        GridGrindService.formulaFor(
            new WorkbookOperation.SetCell(
                "S", "A1", new CellInput.Date(java.time.LocalDate.of(2024, 1, 1))),
            ex));
    assertNull(
        GridGrindService.formulaFor(
            new WorkbookOperation.SetCell(
                "S", "A1", new CellInput.DateTime(java.time.LocalDateTime.of(2024, 1, 1, 0, 0))),
            ex));
  }

  @Test
  void extractsSheetOnlyContextForDeleteSheetOperations() {
    RuntimeException ex = new RuntimeException("test");
    WorkbookOperation deleteSheet = new WorkbookOperation.DeleteSheet("Archive");

    assertNull(GridGrindService.formulaFor(deleteSheet, ex));
    assertEquals("Archive", GridGrindService.sheetNameFor(deleteSheet, ex));
    assertNull(GridGrindService.addressFor(deleteSheet, ex));
    assertNull(GridGrindService.rangeFor(deleteSheet, ex));
  }

  @Test
  void extractsContextForStructuralLayoutOperations() {
    RuntimeException ex = new RuntimeException("test");
    WorkbookOperation mergeCells = new WorkbookOperation.MergeCells("Budget", "A1:B2");
    WorkbookOperation unmergeCells = new WorkbookOperation.UnmergeCells("Budget", "A1:B2");
    WorkbookOperation setColumnWidth = new WorkbookOperation.SetColumnWidth("Budget", 0, 1, 16.0);
    WorkbookOperation setRowHeight = new WorkbookOperation.SetRowHeight("Budget", 0, 1, 28.5);
    WorkbookOperation freezePanes = new WorkbookOperation.FreezePanes("Budget", 1, 1, 1, 1);

    assertNull(GridGrindService.formulaFor(mergeCells, ex));
    assertEquals("Budget", GridGrindService.sheetNameFor(mergeCells, ex));
    assertNull(GridGrindService.addressFor(mergeCells, ex));
    assertEquals("A1:B2", GridGrindService.rangeFor(mergeCells, ex));

    assertNull(GridGrindService.formulaFor(unmergeCells, ex));
    assertEquals("Budget", GridGrindService.sheetNameFor(unmergeCells, ex));
    assertNull(GridGrindService.addressFor(unmergeCells, ex));
    assertEquals("A1:B2", GridGrindService.rangeFor(unmergeCells, ex));

    assertNull(GridGrindService.formulaFor(setColumnWidth, ex));
    assertEquals("Budget", GridGrindService.sheetNameFor(setColumnWidth, ex));
    assertNull(GridGrindService.addressFor(setColumnWidth, ex));
    assertNull(GridGrindService.rangeFor(setColumnWidth, ex));

    assertNull(GridGrindService.formulaFor(setRowHeight, ex));
    assertEquals("Budget", GridGrindService.sheetNameFor(setRowHeight, ex));
    assertNull(GridGrindService.addressFor(setRowHeight, ex));
    assertNull(GridGrindService.rangeFor(setRowHeight, ex));

    assertNull(GridGrindService.formulaFor(freezePanes, ex));
    assertEquals("Budget", GridGrindService.sheetNameFor(freezePanes, ex));
    assertNull(GridGrindService.addressFor(freezePanes, ex));
    assertNull(GridGrindService.rangeFor(freezePanes, ex));
  }
}
