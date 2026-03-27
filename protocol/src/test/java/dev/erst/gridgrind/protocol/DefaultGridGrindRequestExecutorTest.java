package dev.erst.gridgrind.protocol;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.ExcelBorderStyle;
import dev.erst.gridgrind.excel.ExcelCellMetadataSnapshot;
import dev.erst.gridgrind.excel.ExcelCellSnapshot;
import dev.erst.gridgrind.excel.ExcelCellStyleSnapshot;
import dev.erst.gridgrind.excel.ExcelCellValue;
import dev.erst.gridgrind.excel.ExcelComment;
import dev.erst.gridgrind.excel.ExcelFontHeight;
import dev.erst.gridgrind.excel.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.ExcelHyperlink;
import dev.erst.gridgrind.excel.ExcelHyperlinkType;
import dev.erst.gridgrind.excel.ExcelNamedRangeDefinition;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.ExcelNamedRangeSnapshot;
import dev.erst.gridgrind.excel.ExcelNamedRangeTarget;
import dev.erst.gridgrind.excel.ExcelVerticalAlignment;
import dev.erst.gridgrind.excel.ExcelWorkbook;
import dev.erst.gridgrind.excel.InvalidCellAddressException;
import dev.erst.gridgrind.excel.InvalidFormulaException;
import dev.erst.gridgrind.excel.InvalidRangeAddressException;
import dev.erst.gridgrind.excel.NamedRangeNotFoundException;
import dev.erst.gridgrind.excel.SheetNotFoundException;
import dev.erst.gridgrind.excel.UnsupportedFormulaException;
import dev.erst.gridgrind.excel.WorkbookCommand;
import dev.erst.gridgrind.excel.WorkbookCommandExecutor;
import dev.erst.gridgrind.excel.WorkbookNotFoundException;
import dev.erst.gridgrind.excel.WorkbookReadCommand;
import dev.erst.gridgrind.excel.WorkbookReadExecutor;
import dev.erst.gridgrind.excel.XlsxRoundTrip;
import java.io.IOException;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

/** Integration and helper tests for DefaultGridGrindRequestExecutor. */
class DefaultGridGrindRequestExecutorTest {
  @Test
  void executesWorkbookWorkflowAndReturnsOrderedReadResults() throws IOException {
    Path workbookPath = Files.createTempFile("gridgrind-agent-", ".xlsx");
    Files.deleteIfExists(workbookPath);

    GridGrindResponse response =
        new DefaultGridGrindRequestExecutor()
            .execute(
                request(
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
                        new WorkbookOperation.SetCell(
                            "Budget", "B4", new CellInput.Formula("SUM(B2:B3)")),
                        new WorkbookOperation.AutoSizeColumns("Budget"),
                        new WorkbookOperation.EvaluateFormulas(),
                        new WorkbookOperation.ForceFormulaRecalculationOnOpen()),
                    new WorkbookReadOperation.GetWorkbookSummary("workbook"),
                    new WorkbookReadOperation.GetCells(
                        "cells", "Budget", List.of("A1", "B4", "C2")),
                    new WorkbookReadOperation.GetWindow("window", "Budget", "A1", 4, 3)));

    GridGrindResponse.Success success = success(response);
    GridGrindResponse.WorkbookSummary workbook =
        read(success, "workbook", WorkbookReadResult.WorkbookSummaryResult.class).workbook();
    WorkbookReadResult.CellsResult cells =
        read(success, "cells", WorkbookReadResult.CellsResult.class);
    GridGrindResponse.WindowReport window =
        read(success, "window", WorkbookReadResult.WindowResult.class).window();

    assertEquals(GridGrindProtocolVersion.V1, success.protocolVersion());
    assertEquals(workbookPath.toAbsolutePath().toString(), savedPath(success));
    assertTrue(Files.exists(workbookPath));
    assertEquals(List.of("workbook", "cells", "window"), requestIds(success));
    assertEquals(1, workbook.sheetCount());
    assertEquals(List.of("Budget"), workbook.sheetNames());
    assertEquals(0, workbook.namedRangeCount());
    assertTrue(workbook.forceFormulaRecalculationOnOpen());

    assertEquals("Budget", cells.sheetName());
    assertEquals(
        "Item",
        cast(GridGrindResponse.CellReport.TextReport.class, cells.cells().get(0)).stringValue());
    GridGrindResponse.CellReport.FormulaReport formulaCell =
        cast(GridGrindResponse.CellReport.FormulaReport.class, cells.cells().get(1));
    assertEquals("SUM(B2:B3)", formulaCell.formula());
    assertEquals(
        61.0,
        cast(GridGrindResponse.CellReport.NumberReport.class, formulaCell.evaluation())
            .numberValue());
    assertTrue(
        cast(GridGrindResponse.CellReport.BooleanReport.class, cells.cells().get(2))
            .booleanValue());

    assertEquals("Budget", window.sheetName());
    assertEquals("A1", window.topLeftAddress());
    assertEquals(4, window.rows().size());
    assertEquals("A1", window.rows().getFirst().cells().getFirst().address());
  }

  @Test
  void opensExistingWorkbookAndOverwritesSource() throws IOException {
    Path workbookPath = Files.createTempFile("gridgrind-existing-", ".xlsx");

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Budget").setCell("A1", ExcelCellValue.text("Before"));
      workbook.save(workbookPath);
    }

    GridGrindResponse response =
        new DefaultGridGrindRequestExecutor()
            .execute(
                request(
                    new GridGrindRequest.WorkbookSource.ExistingFile(workbookPath.toString()),
                    new GridGrindRequest.WorkbookPersistence.OverwriteSource(),
                    List.of(
                        new WorkbookOperation.SetCell("Budget", "A1", new CellInput.Text("After")),
                        new WorkbookOperation.SetCell("Budget", "B1", new CellInput.Numeric(12.0))),
                    new WorkbookReadOperation.GetCells("cells", "Budget", List.of("A1", "B1"))));

    GridGrindResponse.Success success = success(response);
    WorkbookReadResult.CellsResult cells =
        read(success, "cells", WorkbookReadResult.CellsResult.class);

    assertEquals(workbookPath.toAbsolutePath().toString(), savedPath(success));
    assertEquals(
        "After",
        cast(GridGrindResponse.CellReport.TextReport.class, cells.cells().get(0)).stringValue());
    assertEquals(
        12.0,
        cast(GridGrindResponse.CellReport.NumberReport.class, cells.cells().get(1)).numberValue());
  }

  @Test
  void returnsStructuralReadResultsAndPersistsWorkbookShape() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-layout-ops-");

    GridGrindResponse response =
        new DefaultGridGrindRequestExecutor()
            .execute(
                request(
                    new GridGrindRequest.WorkbookSource.New(),
                    new GridGrindRequest.WorkbookPersistence.SaveAs(workbookPath.toString()),
                    List.of(
                        new WorkbookOperation.EnsureSheet("Budget"),
                        new WorkbookOperation.SetCell(
                            "Budget", "A1", new CellInput.Text("Quarterly")),
                        new WorkbookOperation.MergeCells("Budget", "A1:B1"),
                        new WorkbookOperation.SetColumnWidth("Budget", 0, 1, 16.0),
                        new WorkbookOperation.SetRowHeight("Budget", 0, 0, 28.5),
                        new WorkbookOperation.FreezePanes("Budget", 1, 1, 1, 1)),
                    new WorkbookReadOperation.GetCells("cells", "Budget", List.of("A1")),
                    new WorkbookReadOperation.GetMergedRegions("merged", "Budget"),
                    new WorkbookReadOperation.GetSheetLayout("layout", "Budget"),
                    new WorkbookReadOperation.GetWorkbookSummary("workbook")));

    GridGrindResponse.Success success = success(response);
    WorkbookReadResult.CellsResult cells =
        read(success, "cells", WorkbookReadResult.CellsResult.class);
    WorkbookReadResult.MergedRegionsResult merged =
        read(success, "merged", WorkbookReadResult.MergedRegionsResult.class);
    GridGrindResponse.SheetLayoutReport layout =
        read(success, "layout", WorkbookReadResult.SheetLayoutResult.class).layout();
    GridGrindResponse.WorkbookSummary workbook =
        read(success, "workbook", WorkbookReadResult.WorkbookSummaryResult.class).workbook();

    assertEquals(workbookPath.toAbsolutePath().toString(), savedPath(success));
    assertEquals(
        "Quarterly",
        cast(GridGrindResponse.CellReport.TextReport.class, cells.cells().getFirst())
            .stringValue());
    assertEquals(
        List.of("A1:B1"),
        merged.mergedRegions().stream().map(GridGrindResponse.MergedRegionReport::range).toList());
    assertInstanceOf(GridGrindResponse.FreezePaneReport.Frozen.class, layout.freezePanes());
    GridGrindResponse.FreezePaneReport.Frozen frozen =
        cast(GridGrindResponse.FreezePaneReport.Frozen.class, layout.freezePanes());
    assertEquals(1, frozen.splitColumn());
    assertEquals(1, frozen.splitRow());
    assertEquals(16.0, layout.columns().getFirst().widthCharacters());
    assertEquals(28.5, layout.rows().getFirst().heightPoints());
    assertEquals(List.of("Budget"), workbook.sheetNames());

    assertEquals(List.of("A1:B1"), XlsxRoundTrip.mergedRegions(workbookPath, "Budget"));
    assertEquals(4096, XlsxRoundTrip.columnWidth(workbookPath, "Budget", 0));
    assertEquals((short) 570, XlsxRoundTrip.rowHeightTwips(workbookPath, "Budget", 0));
    assertEquals(
        new XlsxRoundTrip.FreezePaneState.Frozen(1, 1, 1, 1),
        XlsxRoundTrip.freezePaneState(workbookPath, "Budget"));
  }

  @Test
  void returnsAuthoringMetadataAndNamedRangeReadResults() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-authoring-ops-");

    GridGrindResponse response =
        new DefaultGridGrindRequestExecutor()
            .execute(
                request(
                    new GridGrindRequest.WorkbookSource.New(),
                    new GridGrindRequest.WorkbookPersistence.SaveAs(workbookPath.toString()),
                    List.of(
                        new WorkbookOperation.EnsureSheet("Budget"),
                        new WorkbookOperation.SetCell("Budget", "A1", new CellInput.Text("Report")),
                        new WorkbookOperation.SetCell("Budget", "B4", new CellInput.Numeric(61.0)),
                        new WorkbookOperation.SetHyperlink(
                            "Budget", "A1", new HyperlinkTarget.Url("https://example.com/report")),
                        new WorkbookOperation.SetComment(
                            "Budget", "A1", new CommentInput("Review", "GridGrind", true)),
                        new WorkbookOperation.SetNamedRange(
                            "BudgetTotal",
                            new NamedRangeScope.Workbook(),
                            new NamedRangeTarget("Budget", "B4")),
                        new WorkbookOperation.SetNamedRange(
                            "LocalItem",
                            new NamedRangeScope.Sheet("Budget"),
                            new NamedRangeTarget("Budget", "A1:B2"))),
                    new WorkbookReadOperation.GetCells("cells", "Budget", List.of("A1", "B4")),
                    new WorkbookReadOperation.GetHyperlinks(
                        "hyperlinks", "Budget", new CellSelection.Selected(List.of("A1"))),
                    new WorkbookReadOperation.GetComments(
                        "comments", "Budget", new CellSelection.Selected(List.of("A1"))),
                    new WorkbookReadOperation.GetNamedRanges(
                        "ranges", new NamedRangeSelection.All())));

    GridGrindResponse.Success success = success(response);
    GridGrindResponse.CellReport.TextReport linkedCell =
        cast(
            GridGrindResponse.CellReport.TextReport.class,
            read(success, "cells", WorkbookReadResult.CellsResult.class).cells().getFirst());
    WorkbookReadResult.HyperlinksResult hyperlinks =
        read(success, "hyperlinks", WorkbookReadResult.HyperlinksResult.class);
    WorkbookReadResult.CommentsResult comments =
        read(success, "comments", WorkbookReadResult.CommentsResult.class);
    WorkbookReadResult.NamedRangesResult ranges =
        read(success, "ranges", WorkbookReadResult.NamedRangesResult.class);

    assertEquals(ExcelHyperlinkType.URL, linkedCell.hyperlink().type());
    assertEquals("https://example.com/report", linkedCell.hyperlink().target());
    assertEquals("Review", linkedCell.comment().text());
    assertEquals("A1", hyperlinks.hyperlinks().getFirst().address());
    assertEquals(
        "https://example.com/report", hyperlinks.hyperlinks().getFirst().hyperlink().target());
    assertEquals("A1", comments.comments().getFirst().address());
    assertEquals("Review", comments.comments().getFirst().comment().text());
    assertEquals(2, ranges.namedRanges().size());
    assertTrue(
        ranges
            .namedRanges()
            .contains(
                new GridGrindResponse.NamedRangeReport.RangeReport(
                    "BudgetTotal",
                    new NamedRangeScope.Workbook(),
                    "Budget!$B$4",
                    new NamedRangeTarget("Budget", "B4"))));
    assertTrue(
        ranges
            .namedRanges()
            .contains(
                new GridGrindResponse.NamedRangeReport.RangeReport(
                    "LocalItem",
                    new NamedRangeScope.Sheet("Budget"),
                    "Budget!$A$1:Budget!$B$2",
                    new NamedRangeTarget("Budget", "A1:B2"))));

    assertEquals(
        new ExcelHyperlink.Url("https://example.com/report"),
        XlsxRoundTrip.cellMetadata(workbookPath, "Budget", "A1").hyperlink().orElseThrow());
    assertEquals(
        new ExcelComment("Review", "GridGrind", true),
        XlsxRoundTrip.cellMetadata(workbookPath, "Budget", "A1").comment().orElseThrow());
    assertEquals(2, XlsxRoundTrip.namedRanges(workbookPath).size());
  }

  @Test
  void returnsInsightReadResults() {
    GridGrindResponse response =
        new DefaultGridGrindRequestExecutor()
            .execute(
                request(
                    new GridGrindRequest.WorkbookSource.New(),
                    new GridGrindRequest.WorkbookPersistence.None(),
                    List.of(
                        new WorkbookOperation.EnsureSheet("Budget"),
                        new WorkbookOperation.AppendRow(
                            "Budget",
                            List.of(new CellInput.Text("Item"), new CellInput.Text("Amount"))),
                        new WorkbookOperation.AppendRow(
                            "Budget",
                            List.of(new CellInput.Text("Hosting"), new CellInput.Numeric(49.0))),
                        new WorkbookOperation.AppendRow(
                            "Budget",
                            List.of(new CellInput.Text("Domain"), new CellInput.Numeric(12.0))),
                        new WorkbookOperation.SetCell(
                            "Budget", "B4", new CellInput.Formula("SUM(B2:B3)")),
                        new WorkbookOperation.SetNamedRange(
                            "BudgetTotal",
                            new NamedRangeScope.Workbook(),
                            new NamedRangeTarget("Budget", "B4"))),
                    new WorkbookReadOperation.AnalyzeFormulaSurface(
                        "formula", new SheetSelection.All()),
                    new WorkbookReadOperation.AnalyzeSheetSchema("schema", "Budget", "A1", 4, 2),
                    new WorkbookReadOperation.AnalyzeNamedRangeSurface(
                        "ranges", new NamedRangeSelection.All())));

    GridGrindResponse.Success success = success(response);
    GridGrindResponse.FormulaSurfaceReport formula =
        read(success, "formula", WorkbookReadResult.FormulaSurfaceResult.class).analysis();
    GridGrindResponse.SheetSchemaReport schema =
        read(success, "schema", WorkbookReadResult.SheetSchemaResult.class).analysis();
    GridGrindResponse.NamedRangeSurfaceReport ranges =
        read(success, "ranges", WorkbookReadResult.NamedRangeSurfaceResult.class).analysis();

    assertEquals(1, formula.totalFormulaCellCount());
    assertEquals("Budget", formula.sheets().getFirst().sheetName());
    assertEquals("SUM(B2:B3)", formula.sheets().getFirst().formulas().getFirst().formula());
    assertEquals(List.of("B4"), formula.sheets().getFirst().formulas().getFirst().addresses());

    assertEquals("Budget", schema.sheetName());
    assertEquals("A1", schema.topLeftAddress());
    assertEquals(4, schema.rowCount());
    assertEquals(2, schema.columnCount());
    assertEquals(3, schema.dataRowCount());
    assertEquals("Item", schema.columns().getFirst().headerDisplayValue());
    assertEquals("STRING", schema.columns().getFirst().dominantType());

    assertEquals(1, ranges.workbookScopedCount());
    assertEquals(0, ranges.sheetScopedCount());
    assertEquals(1, ranges.rangeBackedCount());
    assertEquals(0, ranges.formulaBackedCount());
    assertEquals("BudgetTotal", ranges.namedRanges().getFirst().name());
    assertEquals(
        GridGrindResponse.NamedRangeBackingKind.RANGE, ranges.namedRanges().getFirst().kind());
  }

  @Test
  void returnsStructuredFailureWhenMoveSheetTargetsMissingSheet() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new GridGrindRequest.WorkbookSource.New(),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        List.of(new WorkbookOperation.MoveSheet("Missing", 0)))));

    assertEquals(GridGrindProblemCode.SHEET_NOT_FOUND, failure.problem().code());
    assertEquals("APPLY_OPERATION", failure.problem().context().stage());
    assertEquals("MOVE_SHEET", failure.problem().context().operationType());
    assertEquals("Missing", failure.problem().context().sheetName());
  }

  @Test
  void returnsStructuredFailureForConflictingRenameTarget() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new GridGrindRequest.WorkbookSource.New(),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        List.of(
                            new WorkbookOperation.EnsureSheet("Budget"),
                            new WorkbookOperation.EnsureSheet("Summary"),
                            new WorkbookOperation.RenameSheet("Budget", "Summary")))));

    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertEquals("APPLY_OPERATION", failure.problem().context().stage());
    assertEquals(2, failure.problem().context().operationIndex());
    assertEquals("RENAME_SHEET", failure.problem().context().operationType());
    assertEquals("Budget", failure.problem().context().sheetName());
    assertEquals("Sheet already exists: Summary", failure.problem().message());
  }

  @Test
  void returnsStructuredFailureForInvalidMoveSheetIndex() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new GridGrindRequest.WorkbookSource.New(),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        List.of(
                            new WorkbookOperation.EnsureSheet("Budget"),
                            new WorkbookOperation.MoveSheet("Budget", 1)))));

    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertEquals("MOVE_SHEET", failure.problem().context().operationType());
    assertEquals("Budget", failure.problem().context().sheetName());
    assertEquals("targetIndex must be between 0 and 0 (inclusive): 1", failure.problem().message());
  }

  @Test
  void returnsStructuredFailureForMergeCellsWithInvalidRange() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new GridGrindRequest.WorkbookSource.New(),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        List.of(
                            new WorkbookOperation.EnsureSheet("Budget"),
                            new WorkbookOperation.MergeCells("Budget", "A1:")))));

    assertEquals(GridGrindProblemCode.INVALID_RANGE_ADDRESS, failure.problem().code());
    assertEquals("MERGE_CELLS", failure.problem().context().operationType());
    assertEquals("Budget", failure.problem().context().sheetName());
    assertEquals("A1:", failure.problem().context().range());
  }

  @Test
  void returnsStructuredFailureForUnmergeCellsWithoutExactMatch() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new GridGrindRequest.WorkbookSource.New(),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        List.of(
                            new WorkbookOperation.EnsureSheet("Budget"),
                            new WorkbookOperation.UnmergeCells("Budget", "A1:B2")))));

    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertEquals("UNMERGE_CELLS", failure.problem().context().operationType());
    assertEquals("Budget", failure.problem().context().sheetName());
    assertEquals("A1:B2", failure.problem().context().range());
    assertEquals("No merged region matches range: A1:B2", failure.problem().message());
  }

  @Test
  void returnsStructuredFailureWhenFreezePanesTargetsMissingSheet() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new GridGrindRequest.WorkbookSource.New(),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        List.of(new WorkbookOperation.FreezePanes("Missing", 1, 1, 1, 1)))));

    assertEquals(GridGrindProblemCode.SHEET_NOT_FOUND, failure.problem().code());
    assertEquals("FREEZE_PANES", failure.problem().context().operationType());
    assertEquals("Missing", failure.problem().context().sheetName());
  }

  @Test
  void returnsStructuredFailureForMissingNamedRangeDuringRead() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new GridGrindRequest.WorkbookSource.New(),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        List.of(new WorkbookOperation.EnsureSheet("Budget")),
                        new WorkbookReadOperation.GetNamedRanges(
                            "ranges",
                            new NamedRangeSelection.Selected(
                                List.of(new NamedRangeSelector.WorkbookScope("BudgetTotal")))))));

    assertEquals(GridGrindProblemCode.NAMED_RANGE_NOT_FOUND, failure.problem().code());
    assertEquals("EXECUTE_READ", failure.problem().context().stage());
    assertEquals(0, failure.problem().context().readIndex());
    assertEquals("GET_NAMED_RANGES", failure.problem().context().readType());
    assertEquals("ranges", failure.problem().context().requestId());
    assertEquals("BudgetTotal", failure.problem().context().namedRangeName());
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

    GridGrindResponse.Success success =
        success(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new GridGrindRequest.WorkbookSource.ExistingFile(workbookPath.toString()),
                        new GridGrindRequest.WorkbookPersistence.OverwriteSource(),
                        List.of(
                            new WorkbookOperation.SetCell(
                                "Budget", "C3", new CellInput.Text("After"))),
                        new WorkbookReadOperation.GetCells("cells", "Budget", List.of("C3")))));

    assertEquals(workbookPath.toAbsolutePath().toString(), savedPath(success));
    assertEquals(
        "After",
        cast(
                GridGrindResponse.CellReport.TextReport.class,
                read(success, "cells", WorkbookReadResult.CellsResult.class).cells().getFirst())
            .stringValue());
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

    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new GridGrindRequest.WorkbookSource.ExistingFile(workbookPath.toString()),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        List.of())));

    assertEquals(GridGrindProblemCode.WORKBOOK_NOT_FOUND, failure.problem().code());
    assertEquals("OPEN_WORKBOOK", failure.problem().context().stage());
    assertEquals(GridGrindProblemCategory.RESOURCE, failure.problem().category());
    assertEquals(GridGrindProblemRecovery.CHANGE_REQUEST, failure.problem().recovery());
    assertEquals(
        workbookPath.toAbsolutePath().toString(), failure.problem().context().sourceWorkbookPath());
  }

  @Test
  void returnsStructuredFailureWithOperationContext() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new GridGrindRequest.WorkbookSource.New(),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        List.of(new WorkbookOperation.AutoSizeColumns("Budget")))));

    assertEquals(GridGrindProblemCode.SHEET_NOT_FOUND, failure.problem().code());
    assertEquals("APPLY_OPERATION", failure.problem().context().stage());
    assertEquals(0, failure.problem().context().operationIndex());
    assertEquals("AUTO_SIZE_COLUMNS", failure.problem().context().operationType());
    assertEquals("Budget", failure.problem().context().sheetName());
  }

  @Test
  void returnsStructuredFailureWithReadContext() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new GridGrindRequest.WorkbookSource.New(),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        List.of(new WorkbookOperation.EnsureSheet("Budget")),
                        new WorkbookReadOperation.GetCells("cells", "Budget", List.of("A1")))));

    assertEquals(GridGrindProblemCode.CELL_NOT_FOUND, failure.problem().code());
    assertEquals("EXECUTE_READ", failure.problem().context().stage());
    assertEquals("GET_CELLS", failure.problem().context().readType());
    assertEquals("cells", failure.problem().context().requestId());
    assertEquals("Budget", failure.problem().context().sheetName());
    assertEquals("A1", failure.problem().context().address());
  }

  @Test
  void returnsStructuredFailureWhenReadTargetsMissingSheet() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new GridGrindRequest.WorkbookSource.New(),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        List.of(),
                        new WorkbookReadOperation.GetSheetSummary("sheet", "Missing"))));

    assertEquals(GridGrindProblemCode.SHEET_NOT_FOUND, failure.problem().code());
    assertEquals("EXECUTE_READ", failure.problem().context().stage());
    assertEquals("GET_SHEET_SUMMARY", failure.problem().context().readType());
    assertEquals("sheet", failure.problem().context().requestId());
    assertEquals("Missing", failure.problem().context().sheetName());
  }

  @Test
  void returnsStructuredFailureForInvalidPersistTarget() throws IOException {
    Path parentFile = Files.createTempFile("gridgrind-persist-target-", ".tmp");
    Path workbookPath = parentFile.resolve("book.xlsx");

    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new GridGrindRequest.WorkbookSource.New(),
                        new GridGrindRequest.WorkbookPersistence.SaveAs(workbookPath.toString()),
                        List.of(new WorkbookOperation.EnsureSheet("Budget")))));

    assertEquals(GridGrindProblemCode.IO_ERROR, failure.problem().code());
    assertEquals("PERSIST_WORKBOOK", failure.problem().context().stage());
    assertEquals(
        workbookPath.toAbsolutePath().toString(), failure.problem().context().persistencePath());
  }

  @Test
  void doesNotPersistWorkbookWhenReadFails() throws IOException {
    Path workbookPath = Files.createTempFile("gridgrind-analysis-failure-", ".xlsx");
    Files.deleteIfExists(workbookPath);

    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new GridGrindRequest.WorkbookSource.New(),
                        new GridGrindRequest.WorkbookPersistence.SaveAs(workbookPath.toString()),
                        List.of(new WorkbookOperation.EnsureSheet("Budget")),
                        new WorkbookReadOperation.GetCells("cells", "Budget", List.of("A1")))));

    assertEquals(GridGrindProblemCode.CELL_NOT_FOUND, failure.problem().code());
    assertFalse(Files.exists(workbookPath));
  }

  @Test
  void returnsStructuredFailureForInvalidOverwriteSourceUsage() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new GridGrindRequest.WorkbookSource.New(),
                        new GridGrindRequest.WorkbookPersistence.OverwriteSource(),
                        List.of())));

    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertEquals("VALIDATE_REQUEST", failure.problem().context().stage());
  }

  @Test
  void returnsFormulaErrorForInvalidFormulaOperations() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new GridGrindRequest.WorkbookSource.New(),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        List.of(
                            new WorkbookOperation.SetCell(
                                "Data", "A1", new CellInput.Formula("SUM(")),
                            new WorkbookOperation.EvaluateFormulas()))));

    assertEquals(GridGrindProblemCode.INVALID_FORMULA, failure.problem().code());
    assertEquals(GridGrindProblemCategory.FORMULA, failure.problem().category());
    assertEquals("APPLY_OPERATION", failure.problem().context().stage());
    assertEquals("Invalid formula at Data!A1: SUM(", failure.problem().message());
    assertEquals("A1", failure.problem().context().address());
    assertEquals("SUM(", failure.problem().context().formula());
  }

  @Test
  void returnsFormulaErrorForMalformedParserStateFormulaOperations() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new GridGrindRequest.WorkbookSource.New(),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        List.of(
                            new WorkbookOperation.SetCell(
                                "Data", "A1", new CellInput.Formula("[^owe_e`ffffff"))))));

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
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new GridGrindRequest.WorkbookSource.New(),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        List.of(
                            new WorkbookOperation.SetCell("Data", "A1", new CellInput.Numeric(1.0)),
                            new WorkbookOperation.SetCell("Data", "B1", new CellInput.Numeric(2.0)),
                            new WorkbookOperation.SetCell(
                                "Data", "C1", new CellInput.Formula("TEXTAFTER(\"a,b\",\",\")")),
                            new WorkbookOperation.EvaluateFormulas()))));

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
    DefaultGridGrindRequestExecutor executor =
        new DefaultGridGrindRequestExecutor(
            new WorkbookCommandExecutor(),
            new WorkbookReadExecutor(),
            workbook -> {
              throw new IOException("close failed");
            });

    GridGrindResponse.Failure failure =
        failure(
            executor.execute(
                request(
                    new GridGrindRequest.WorkbookSource.New(),
                    new GridGrindRequest.WorkbookPersistence.None(),
                    List.of(new WorkbookOperation.EnsureSheet("Budget")))));

    assertEquals(GridGrindProblemCode.IO_ERROR, failure.problem().code());
    assertEquals("EXECUTE_REQUEST", failure.problem().context().stage());
    assertEquals("close failed", failure.problem().message());
    assertEquals(GridGrindProblemRecovery.CHECK_ENVIRONMENT, failure.problem().recovery());
  }

  @Test
  void preservesPrimaryFailureWhenWorkbookCloseAlsoFails() {
    DefaultGridGrindRequestExecutor executor =
        new DefaultGridGrindRequestExecutor(
            new WorkbookCommandExecutor(),
            new WorkbookReadExecutor(),
            workbook -> {
              throw new IOException("close failed");
            });

    GridGrindResponse.Failure failure =
        failure(
            executor.execute(
                request(
                    new GridGrindRequest.WorkbookSource.New(),
                    new GridGrindRequest.WorkbookPersistence.None(),
                    List.of(new WorkbookOperation.AutoSizeColumns("Budget")))));

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
    GridGrindResponse.Failure failure =
        failure(new DefaultGridGrindRequestExecutor().execute(null));

    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertEquals("VALIDATE_REQUEST", failure.problem().context().stage());
  }

  @Test
  void classifiesProblemCodesAndMessagesDeterministically() {
    assertEquals(
        GridGrindProblemCode.INVALID_CELL_ADDRESS,
        DefaultGridGrindRequestExecutor.problemCodeFor(
            new InvalidCellAddressException("BAD!", new IllegalArgumentException("bad"))));
    assertEquals(
        GridGrindProblemCode.INVALID_RANGE_ADDRESS,
        DefaultGridGrindRequestExecutor.problemCodeFor(
            new InvalidRangeAddressException("A1:", new IllegalArgumentException("bad"))));
    assertEquals(
        GridGrindProblemCode.INVALID_FORMULA,
        DefaultGridGrindRequestExecutor.problemCodeFor(
            new InvalidFormulaException(
                "Budget", "B4", "SUM(", "bad formula", new IllegalArgumentException("bad"))));
    assertEquals(
        GridGrindProblemCode.UNSUPPORTED_FORMULA,
        DefaultGridGrindRequestExecutor.problemCodeFor(
            new UnsupportedFormulaException(
                "Budget",
                "C1",
                "LAMBDA(x,x+1)(2)",
                "unsupported",
                new IllegalArgumentException("bad"))));
    assertEquals(
        GridGrindProblemCode.WORKBOOK_NOT_FOUND,
        DefaultGridGrindRequestExecutor.problemCodeFor(
            new WorkbookNotFoundException(Path.of("missing.xlsx"))));
    assertEquals(
        GridGrindProblemCode.SHEET_NOT_FOUND,
        DefaultGridGrindRequestExecutor.problemCodeFor(new SheetNotFoundException("Budget")));
    assertEquals(
        GridGrindProblemCode.IO_ERROR,
        DefaultGridGrindRequestExecutor.problemCodeFor(new IOException("disk")));
    assertEquals(
        GridGrindProblemCode.INVALID_REQUEST,
        DefaultGridGrindRequestExecutor.problemCodeFor(
            new IllegalArgumentException("bad request")));
    assertEquals(
        GridGrindProblemCode.INTERNAL_ERROR,
        DefaultGridGrindRequestExecutor.problemCodeFor(new UnsupportedOperationException()));
    assertEquals(
        "bad request",
        DefaultGridGrindRequestExecutor.messageFor(new IllegalArgumentException("bad request")));
    assertEquals(
        "UnsupportedOperationException",
        DefaultGridGrindRequestExecutor.messageFor(new UnsupportedOperationException()));
  }

  @Test
  void executesRangeAndStyleWorkflowAndSurfacesStyledCells() throws IOException {
    Path workbookPath = Files.createTempFile("gridgrind-range-style-", ".xlsx");
    Files.deleteIfExists(workbookPath);

    GridGrindResponse.Success success =
        success(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new GridGrindRequest.WorkbookSource.New(),
                        new GridGrindRequest.WorkbookPersistence.SaveAs(workbookPath.toString()),
                        List.of(
                            new WorkbookOperation.EnsureSheet("Budget"),
                            new WorkbookOperation.SetRange(
                                "Budget",
                                "A1:B2",
                                List.of(
                                    List.of(
                                        new CellInput.Text("Item"), new CellInput.Text("Amount")),
                                    List.of(
                                        new CellInput.Text("Hosting"),
                                        new CellInput.Numeric(49.0)))),
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
                        new WorkbookReadOperation.GetCells(
                            "cells", "Budget", List.of("A1", "A2", "B3", "C1")),
                        new WorkbookReadOperation.GetWindow("window", "Budget", "A1", 3, 3))));

    WorkbookReadResult.CellsResult cells =
        read(success, "cells", WorkbookReadResult.CellsResult.class);
    GridGrindResponse.WindowReport window =
        read(success, "window", WorkbookReadResult.WindowResult.class).window();

    assertTrue(Files.exists(workbookPath));
    assertEquals(
        ExcelHorizontalAlignment.CENTER, cells.cells().getFirst().style().horizontalAlignment());
    assertTrue(cells.cells().getFirst().style().bold());
    assertTrue(cells.cells().getFirst().style().wrapText());
    assertEquals("BLANK", cells.cells().get(1).effectiveType());
    assertEquals("49", cells.cells().get(2).displayValue());
    assertTrue(
        window.rows().getFirst().cells().stream().anyMatch(cell -> "C1".equals(cell.address())));
    assertTrue(cells.cells().get(3).style().italic());
  }

  @Test
  void executesFormattingDepthWorkflowAndPersistsReportedStyleState() throws IOException {
    Path workbookPath = Files.createTempFile("gridgrind-format-depth-", ".xlsx");
    Files.deleteIfExists(workbookPath);

    GridGrindResponse.Success success =
        success(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new GridGrindRequest.WorkbookSource.New(),
                        new GridGrindRequest.WorkbookPersistence.SaveAs(workbookPath.toString()),
                        List.of(
                            new WorkbookOperation.EnsureSheet("Budget"),
                            new WorkbookOperation.SetCell(
                                "Budget", "A1", new CellInput.Text("Item")),
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
                        new WorkbookReadOperation.GetCells("cells", "Budget", List.of("A1")))));

    GridGrindResponse.CellStyleReport style =
        read(success, "cells", WorkbookReadResult.CellsResult.class).cells().getFirst().style();

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
    GridGrindResponse.Success success =
        success(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new GridGrindRequest.WorkbookSource.New(),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        List.of(
                            new WorkbookOperation.EnsureSheet("Data"),
                            new WorkbookOperation.SetCell(
                                "Data", "A1", new CellInput.Formula("1/0")),
                            new WorkbookOperation.EvaluateFormulas()),
                        new WorkbookReadOperation.GetCells("cells", "Data", List.of("A1")))));

    GridGrindResponse.CellReport cell =
        read(success, "cells", WorkbookReadResult.CellsResult.class).cells().getFirst();
    assertInstanceOf(GridGrindResponse.CellReport.FormulaReport.class, cell);
    GridGrindResponse.CellReport evaluation =
        cast(GridGrindResponse.CellReport.FormulaReport.class, cell).evaluation();
    assertInstanceOf(GridGrindResponse.CellReport.ErrorReport.class, evaluation);
    assertEquals("ERROR", evaluation.effectiveType());
  }

  @Test
  void extractsFormulaFromSetCellOperationWhenExceptionCarriesNone() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new GridGrindRequest.WorkbookSource.New(),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        List.of(
                            new WorkbookOperation.EnsureSheet("Data"),
                            new WorkbookOperation.SetCell(
                                "Data", "INVALID!", new CellInput.Formula("SUM(B1:B2)"))))));

    assertEquals("APPLY_OPERATION", failure.problem().context().stage());
    assertEquals("SUM(B1:B2)", failure.problem().context().formula());
  }

  @Test
  void persistencePathResolvesCorrectlyForAllPersistenceAndSourceCombinations() {
    DefaultGridGrindRequestExecutor executor = new DefaultGridGrindRequestExecutor();

    GridGrindRequest.WorkbookSource newSource = new GridGrindRequest.WorkbookSource.New();
    GridGrindRequest.WorkbookSource existingFile =
        new GridGrindRequest.WorkbookSource.ExistingFile("/tmp/source.xlsx");
    GridGrindRequest.WorkbookPersistence none = new GridGrindRequest.WorkbookPersistence.None();
    GridGrindRequest.WorkbookPersistence overwrite =
        new GridGrindRequest.WorkbookPersistence.OverwriteSource();
    GridGrindRequest.WorkbookPersistence saveAs =
        new GridGrindRequest.WorkbookPersistence.SaveAs("/tmp/out.xlsx");

    assertEquals(
        Path.of("/tmp/out.xlsx").toAbsolutePath().toString(),
        executor.persistencePath(newSource, saveAs));
    assertEquals(
        Path.of("/tmp/source.xlsx").toAbsolutePath().toString(),
        executor.persistencePath(existingFile, overwrite));
    assertNull(executor.persistencePath(newSource, overwrite));
    assertNull(executor.persistencePath(newSource, none));
    assertNull(executor.persistencePath(existingFile, none));
  }

  @Test
  void guardsCatastrophicRuntimeExceptionsAndProducesExecuteRequestFailure() {
    int[] callCount = {0};
    DefaultGridGrindRequestExecutor executor =
        new DefaultGridGrindRequestExecutor(
            new WorkbookCommandExecutor(),
            new WorkbookReadExecutor(),
            workbook -> {
              int count = callCount[0];
              callCount[0] = count + 1;
              if (count == 0) {
                throw new IllegalStateException("catastrophic close failure");
              }
            });

    GridGrindResponse.Failure failure =
        failure(
            executor.execute(
                request(
                    new GridGrindRequest.WorkbookSource.New(),
                    new GridGrindRequest.WorkbookPersistence.None(),
                    List.of(new WorkbookOperation.EnsureSheet("Budget")))));

    assertEquals(GridGrindProblemCode.INTERNAL_ERROR, failure.problem().code());
    assertEquals("EXECUTE_REQUEST", failure.problem().context().stage());
  }

  @Test
  void returnsStructuredFailureForInvalidRangeOperations() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new GridGrindRequest.WorkbookSource.New(),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        List.of(
                            new WorkbookOperation.EnsureSheet("Budget"),
                            new WorkbookOperation.ClearRange("Budget", "A1:")))));

    assertEquals(GridGrindProblemCode.INVALID_RANGE_ADDRESS, failure.problem().code());
    assertEquals("CLEAR_RANGE", failure.problem().context().operationType());
    assertEquals("Budget", failure.problem().context().sheetName());
    assertEquals("A1:", failure.problem().context().range());
  }

  @Test
  void returnsStructuredFailureForSetRangeWithInvalidRange() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new GridGrindRequest.WorkbookSource.New(),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        List.of(
                            new WorkbookOperation.EnsureSheet("Budget"),
                            new WorkbookOperation.SetRange(
                                "Budget", "A1:", List.of(List.of(new CellInput.Text("x"))))))));

    assertEquals("APPLY_OPERATION", failure.problem().context().stage());
    assertEquals("SET_RANGE", failure.problem().context().operationType());
    assertEquals("Budget", failure.problem().context().sheetName());
    assertEquals("A1:", failure.problem().context().range());
  }

  @Test
  void returnsStructuredFailureForApplyStyleWithInvalidRange() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new GridGrindRequest.WorkbookSource.New(),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        List.of(
                            new WorkbookOperation.EnsureSheet("Budget"),
                            new WorkbookOperation.ApplyStyle(
                                "Budget",
                                "A1:",
                                new CellStyleInput(
                                    null, true, null, null, null, null, null, null, null, null,
                                    null, null, null))))));

    assertEquals("APPLY_OPERATION", failure.problem().context().stage());
    assertEquals("APPLY_STYLE", failure.problem().context().operationType());
    assertEquals("Budget", failure.problem().context().sheetName());
    assertEquals("A1:", failure.problem().context().range());
  }

  @Test
  void returnsStructuredFailureForAppendRowWithInvalidFormula() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new GridGrindRequest.WorkbookSource.New(),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        List.of(
                            new WorkbookOperation.EnsureSheet("Budget"),
                            new WorkbookOperation.AppendRow(
                                "Budget", List.of(new CellInput.Formula("SUM(")))))));

    assertEquals("APPLY_OPERATION", failure.problem().context().stage());
    assertEquals("APPEND_ROW", failure.problem().context().operationType());
    assertEquals("Budget", failure.problem().context().sheetName());
  }

  @Test
  void returnsStructuredFailureForEnsureSheetWithInvalidSheetName() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new GridGrindRequest.WorkbookSource.New(),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        List.of(new WorkbookOperation.EnsureSheet("[Budget]")))));

    assertEquals("APPLY_OPERATION", failure.problem().context().stage());
    assertEquals("ENSURE_SHEET", failure.problem().context().operationType());
    assertEquals("[Budget]", failure.problem().context().sheetName());
  }

  @Test
  void extractsNullContextForOperationsWithNoSheetAddressRangeOrFormula() {
    RuntimeException exception = new RuntimeException("test");
    WorkbookOperation forceRecalc = new WorkbookOperation.ForceFormulaRecalculationOnOpen();
    WorkbookOperation evalFormulas = new WorkbookOperation.EvaluateFormulas();
    WorkbookOperation appendRow =
        new WorkbookOperation.AppendRow("Budget", List.of(new CellInput.Text("x")));
    WorkbookOperation ensureSheet = new WorkbookOperation.EnsureSheet("Budget");

    assertNull(DefaultGridGrindRequestExecutor.formulaFor(forceRecalc, exception));
    assertNull(DefaultGridGrindRequestExecutor.formulaFor(evalFormulas, exception));
    assertNull(DefaultGridGrindRequestExecutor.formulaFor(appendRow, exception));
    assertNull(DefaultGridGrindRequestExecutor.formulaFor(ensureSheet, exception));
    assertNull(DefaultGridGrindRequestExecutor.sheetNameFor(forceRecalc, exception));
    assertNull(DefaultGridGrindRequestExecutor.sheetNameFor(evalFormulas, exception));
    assertNull(DefaultGridGrindRequestExecutor.addressFor(forceRecalc, exception));
    assertNull(DefaultGridGrindRequestExecutor.addressFor(evalFormulas, exception));
    assertNull(DefaultGridGrindRequestExecutor.rangeFor(forceRecalc, exception));
    assertNull(DefaultGridGrindRequestExecutor.rangeFor(evalFormulas, exception));
  }

  @Test
  void extractsNullFormulaFromSetCellWithNonFormulaValueWhenExceptionCarriesNone() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new GridGrindRequest.WorkbookSource.New(),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        List.of(
                            new WorkbookOperation.EnsureSheet("Budget"),
                            new WorkbookOperation.SetCell(
                                "Budget", "INVALID!", new CellInput.Text("hello"))))));

    assertEquals("APPLY_OPERATION", failure.problem().context().stage());
    assertEquals("SET_CELL", failure.problem().context().operationType());
    assertNull(failure.problem().context().formula());
  }

  @Test
  void formulaForSetCellReturnsNullForAllNonFormulaValueTypes() {
    RuntimeException exception = new RuntimeException("test");

    assertNull(
        DefaultGridGrindRequestExecutor.formulaFor(
            new WorkbookOperation.SetCell("S", "A1", new CellInput.Blank()), exception));
    assertNull(
        DefaultGridGrindRequestExecutor.formulaFor(
            new WorkbookOperation.SetCell("S", "A1", new CellInput.Numeric(1.0)), exception));
    assertNull(
        DefaultGridGrindRequestExecutor.formulaFor(
            new WorkbookOperation.SetCell("S", "A1", new CellInput.BooleanValue(true)), exception));
    assertNull(
        DefaultGridGrindRequestExecutor.formulaFor(
            new WorkbookOperation.SetCell("S", "A1", new CellInput.Date(LocalDate.of(2024, 1, 1))),
            exception));
    assertNull(
        DefaultGridGrindRequestExecutor.formulaFor(
            new WorkbookOperation.SetCell(
                "S", "A1", new CellInput.DateTime(LocalDateTime.of(2024, 1, 1, 0, 0))),
            exception));
  }

  @Test
  void extractsSheetOnlyContextForDeleteSheetOperations() {
    RuntimeException exception = new RuntimeException("test");
    WorkbookOperation deleteSheet = new WorkbookOperation.DeleteSheet("Archive");

    assertNull(DefaultGridGrindRequestExecutor.formulaFor(deleteSheet, exception));
    assertEquals("Archive", DefaultGridGrindRequestExecutor.sheetNameFor(deleteSheet, exception));
    assertNull(DefaultGridGrindRequestExecutor.addressFor(deleteSheet, exception));
    assertNull(DefaultGridGrindRequestExecutor.rangeFor(deleteSheet, exception));
  }

  @Test
  void extractsContextForStructuralLayoutOperations() {
    RuntimeException exception = new RuntimeException("test");
    WorkbookOperation mergeCells = new WorkbookOperation.MergeCells("Budget", "A1:B2");
    WorkbookOperation unmergeCells = new WorkbookOperation.UnmergeCells("Budget", "A1:B2");
    WorkbookOperation setColumnWidth = new WorkbookOperation.SetColumnWidth("Budget", 0, 1, 16.0);
    WorkbookOperation setRowHeight = new WorkbookOperation.SetRowHeight("Budget", 0, 1, 28.5);
    WorkbookOperation freezePanes = new WorkbookOperation.FreezePanes("Budget", 1, 1, 1, 1);

    assertNull(DefaultGridGrindRequestExecutor.formulaFor(mergeCells, exception));
    assertEquals("Budget", DefaultGridGrindRequestExecutor.sheetNameFor(mergeCells, exception));
    assertNull(DefaultGridGrindRequestExecutor.addressFor(mergeCells, exception));
    assertEquals("A1:B2", DefaultGridGrindRequestExecutor.rangeFor(mergeCells, exception));

    assertNull(DefaultGridGrindRequestExecutor.formulaFor(unmergeCells, exception));
    assertEquals("Budget", DefaultGridGrindRequestExecutor.sheetNameFor(unmergeCells, exception));
    assertNull(DefaultGridGrindRequestExecutor.addressFor(unmergeCells, exception));
    assertEquals("A1:B2", DefaultGridGrindRequestExecutor.rangeFor(unmergeCells, exception));

    assertNull(DefaultGridGrindRequestExecutor.formulaFor(setColumnWidth, exception));
    assertEquals("Budget", DefaultGridGrindRequestExecutor.sheetNameFor(setColumnWidth, exception));
    assertNull(DefaultGridGrindRequestExecutor.addressFor(setColumnWidth, exception));
    assertNull(DefaultGridGrindRequestExecutor.rangeFor(setColumnWidth, exception));

    assertNull(DefaultGridGrindRequestExecutor.formulaFor(setRowHeight, exception));
    assertEquals("Budget", DefaultGridGrindRequestExecutor.sheetNameFor(setRowHeight, exception));
    assertNull(DefaultGridGrindRequestExecutor.addressFor(setRowHeight, exception));
    assertNull(DefaultGridGrindRequestExecutor.rangeFor(setRowHeight, exception));

    assertNull(DefaultGridGrindRequestExecutor.formulaFor(freezePanes, exception));
    assertEquals("Budget", DefaultGridGrindRequestExecutor.sheetNameFor(freezePanes, exception));
    assertNull(DefaultGridGrindRequestExecutor.addressFor(freezePanes, exception));
    assertNull(DefaultGridGrindRequestExecutor.rangeFor(freezePanes, exception));
  }

  @Test
  void convertsWaveThreeOperationsIntoWorkbookCommands() {
    WorkbookCommand setHyperlink =
        DefaultGridGrindRequestExecutor.toCommand(
            new WorkbookOperation.SetHyperlink(
                "Budget", "A1", new HyperlinkTarget.Url("https://example.com/report")));
    WorkbookCommand clearHyperlink =
        DefaultGridGrindRequestExecutor.toCommand(
            new WorkbookOperation.ClearHyperlink("Budget", "A1"));
    WorkbookCommand setComment =
        DefaultGridGrindRequestExecutor.toCommand(
            new WorkbookOperation.SetComment(
                "Budget", "A1", new CommentInput("Review", "GridGrind", false)));
    WorkbookCommand clearComment =
        DefaultGridGrindRequestExecutor.toCommand(
            new WorkbookOperation.ClearComment("Budget", "A1"));
    WorkbookCommand setNamedRange =
        DefaultGridGrindRequestExecutor.toCommand(
            new WorkbookOperation.SetNamedRange(
                "BudgetTotal",
                new NamedRangeScope.Workbook(),
                new NamedRangeTarget("Budget", "B4")));
    WorkbookCommand deleteNamedRange =
        DefaultGridGrindRequestExecutor.toCommand(
            new WorkbookOperation.DeleteNamedRange(
                "BudgetTotal", new NamedRangeScope.Sheet("Budget")));

    assertInstanceOf(WorkbookCommand.SetHyperlink.class, setHyperlink);
    assertInstanceOf(WorkbookCommand.ClearHyperlink.class, clearHyperlink);
    assertInstanceOf(WorkbookCommand.SetComment.class, setComment);
    assertInstanceOf(WorkbookCommand.ClearComment.class, clearComment);
    assertInstanceOf(WorkbookCommand.SetNamedRange.class, setNamedRange);
    assertInstanceOf(WorkbookCommand.DeleteNamedRange.class, deleteNamedRange);
    assertEquals(
        new ExcelHyperlink.Url("https://example.com/report"),
        cast(WorkbookCommand.SetHyperlink.class, setHyperlink).target());
    assertEquals(
        new ExcelComment("Review", "GridGrind", false),
        cast(WorkbookCommand.SetComment.class, setComment).comment());
    assertEquals(
        new ExcelNamedRangeDefinition(
            "BudgetTotal",
            new ExcelNamedRangeScope.WorkbookScope(),
            new ExcelNamedRangeTarget("Budget", "B4")),
        cast(WorkbookCommand.SetNamedRange.class, setNamedRange).definition());
    assertEquals(
        new ExcelNamedRangeScope.SheetScope("Budget"),
        cast(WorkbookCommand.DeleteNamedRange.class, deleteNamedRange).scope());
  }

  @Test
  void convertsRemainingWorkbookOperationsIntoWorkbookCommands() {
    assertInstanceOf(
        WorkbookCommand.CreateSheet.class,
        DefaultGridGrindRequestExecutor.toCommand(new WorkbookOperation.EnsureSheet("Budget")));
    assertInstanceOf(
        WorkbookCommand.RenameSheet.class,
        DefaultGridGrindRequestExecutor.toCommand(
            new WorkbookOperation.RenameSheet("Budget", "Summary")));
    assertInstanceOf(
        WorkbookCommand.DeleteSheet.class,
        DefaultGridGrindRequestExecutor.toCommand(new WorkbookOperation.DeleteSheet("Budget")));
    assertInstanceOf(
        WorkbookCommand.MoveSheet.class,
        DefaultGridGrindRequestExecutor.toCommand(new WorkbookOperation.MoveSheet("Budget", 0)));
    assertInstanceOf(
        WorkbookCommand.MergeCells.class,
        DefaultGridGrindRequestExecutor.toCommand(
            new WorkbookOperation.MergeCells("Budget", "A1:B2")));
    assertInstanceOf(
        WorkbookCommand.UnmergeCells.class,
        DefaultGridGrindRequestExecutor.toCommand(
            new WorkbookOperation.UnmergeCells("Budget", "A1:B2")));
    assertInstanceOf(
        WorkbookCommand.SetColumnWidth.class,
        DefaultGridGrindRequestExecutor.toCommand(
            new WorkbookOperation.SetColumnWidth("Budget", 0, 1, 16.0)));
    assertInstanceOf(
        WorkbookCommand.SetRowHeight.class,
        DefaultGridGrindRequestExecutor.toCommand(
            new WorkbookOperation.SetRowHeight("Budget", 0, 1, 28.5)));
    assertInstanceOf(
        WorkbookCommand.FreezePanes.class,
        DefaultGridGrindRequestExecutor.toCommand(
            new WorkbookOperation.FreezePanes("Budget", 1, 1, 1, 1)));
    assertInstanceOf(
        WorkbookCommand.SetCell.class,
        DefaultGridGrindRequestExecutor.toCommand(
            new WorkbookOperation.SetCell("Budget", "A1", new CellInput.Text("x"))));
    assertInstanceOf(
        WorkbookCommand.SetRange.class,
        DefaultGridGrindRequestExecutor.toCommand(
            new WorkbookOperation.SetRange(
                "Budget", "A1:B1", List.of(List.of(new CellInput.Text("x"))))));
    assertInstanceOf(
        WorkbookCommand.ClearRange.class,
        DefaultGridGrindRequestExecutor.toCommand(
            new WorkbookOperation.ClearRange("Budget", "A1:B1")));
    assertInstanceOf(
        WorkbookCommand.AppendRow.class,
        DefaultGridGrindRequestExecutor.toCommand(
            new WorkbookOperation.AppendRow("Budget", List.of(new CellInput.Text("x")))));
    assertInstanceOf(
        WorkbookCommand.AutoSizeColumns.class,
        DefaultGridGrindRequestExecutor.toCommand(new WorkbookOperation.AutoSizeColumns("Budget")));
    assertInstanceOf(
        WorkbookCommand.EvaluateAllFormulas.class,
        DefaultGridGrindRequestExecutor.toCommand(new WorkbookOperation.EvaluateFormulas()));
    assertInstanceOf(
        WorkbookCommand.ForceFormulaRecalculationOnOpen.class,
        DefaultGridGrindRequestExecutor.toCommand(
            new WorkbookOperation.ForceFormulaRecalculationOnOpen()));
  }

  @Test
  void convertsReadOperationsIntoWorkbookReadCommands() {
    WorkbookReadCommand workbookSummary =
        DefaultGridGrindRequestExecutor.toReadCommand(
            new WorkbookReadOperation.GetWorkbookSummary("workbook"));
    WorkbookReadCommand namedRanges =
        DefaultGridGrindRequestExecutor.toReadCommand(
            new WorkbookReadOperation.GetNamedRanges(
                "ranges",
                new NamedRangeSelection.Selected(
                    List.of(
                        new NamedRangeSelector.ByName("BudgetTotal"),
                        new NamedRangeSelector.SheetScope("LocalItem", "Budget")))));
    WorkbookReadCommand sheetSummary =
        DefaultGridGrindRequestExecutor.toReadCommand(
            new WorkbookReadOperation.GetSheetSummary("sheet", "Budget"));
    WorkbookReadCommand cells =
        DefaultGridGrindRequestExecutor.toReadCommand(
            new WorkbookReadOperation.GetCells("cells", "Budget", List.of("A1")));
    WorkbookReadCommand window =
        DefaultGridGrindRequestExecutor.toReadCommand(
            new WorkbookReadOperation.GetWindow("window", "Budget", "A1", 2, 2));
    WorkbookReadCommand merged =
        DefaultGridGrindRequestExecutor.toReadCommand(
            new WorkbookReadOperation.GetMergedRegions("merged", "Budget"));
    WorkbookReadCommand hyperlinks =
        DefaultGridGrindRequestExecutor.toReadCommand(
            new WorkbookReadOperation.GetHyperlinks(
                "hyperlinks", "Budget", new CellSelection.AllUsedCells()));
    WorkbookReadCommand comments =
        DefaultGridGrindRequestExecutor.toReadCommand(
            new WorkbookReadOperation.GetComments(
                "comments", "Budget", new CellSelection.Selected(List.of("A1"))));
    WorkbookReadCommand layout =
        DefaultGridGrindRequestExecutor.toReadCommand(
            new WorkbookReadOperation.GetSheetLayout("layout", "Budget"));
    WorkbookReadCommand formulaSurface =
        DefaultGridGrindRequestExecutor.toReadCommand(
            new WorkbookReadOperation.AnalyzeFormulaSurface(
                "formula", new SheetSelection.Selected(List.of("Budget"))));
    WorkbookReadCommand schema =
        DefaultGridGrindRequestExecutor.toReadCommand(
            new WorkbookReadOperation.AnalyzeSheetSchema("schema", "Budget", "A1", 3, 2));
    WorkbookReadCommand namedRangeSurface =
        DefaultGridGrindRequestExecutor.toReadCommand(
            new WorkbookReadOperation.AnalyzeNamedRangeSurface(
                "surface", new NamedRangeSelection.All()));

    assertInstanceOf(WorkbookReadCommand.GetWorkbookSummary.class, workbookSummary);
    assertInstanceOf(WorkbookReadCommand.GetNamedRanges.class, namedRanges);
    assertInstanceOf(WorkbookReadCommand.GetSheetSummary.class, sheetSummary);
    assertInstanceOf(WorkbookReadCommand.GetCells.class, cells);
    assertInstanceOf(WorkbookReadCommand.GetWindow.class, window);
    assertInstanceOf(WorkbookReadCommand.GetMergedRegions.class, merged);
    assertInstanceOf(WorkbookReadCommand.GetHyperlinks.class, hyperlinks);
    assertInstanceOf(WorkbookReadCommand.GetComments.class, comments);
    assertInstanceOf(WorkbookReadCommand.GetSheetLayout.class, layout);
    assertInstanceOf(WorkbookReadCommand.AnalyzeFormulaSurface.class, formulaSurface);
    assertInstanceOf(WorkbookReadCommand.AnalyzeSheetSchema.class, schema);
    assertInstanceOf(WorkbookReadCommand.AnalyzeNamedRangeSurface.class, namedRangeSurface);
    assertInstanceOf(
        dev.erst.gridgrind.excel.ExcelNamedRangeSelection.Selected.class,
        cast(WorkbookReadCommand.GetNamedRanges.class, namedRanges).selection());
    assertInstanceOf(
        dev.erst.gridgrind.excel.ExcelCellSelection.AllUsedCells.class,
        cast(WorkbookReadCommand.GetHyperlinks.class, hyperlinks).selection());
    assertInstanceOf(
        dev.erst.gridgrind.excel.ExcelCellSelection.Selected.class,
        cast(WorkbookReadCommand.GetComments.class, comments).selection());
    assertInstanceOf(
        dev.erst.gridgrind.excel.ExcelSheetSelection.Selected.class,
        cast(WorkbookReadCommand.AnalyzeFormulaSurface.class, formulaSurface).selection());
    assertInstanceOf(
        dev.erst.gridgrind.excel.ExcelNamedRangeSelection.All.class,
        cast(WorkbookReadCommand.AnalyzeNamedRangeSurface.class, namedRangeSurface).selection());
  }

  @Test
  void convertsReadResultsIntoProtocolReadResults() {
    ExcelCellSnapshot.BlankSnapshot blank =
        new ExcelCellSnapshot.BlankSnapshot(
            "A1", "BLANK", "", defaultStyle(), ExcelCellMetadataSnapshot.empty());

    WorkbookReadResult workbookSummary =
        DefaultGridGrindRequestExecutor.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummaryResult(
                "workbook",
                new dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummary(
                    1, List.of("Budget"), 1, true)));
    WorkbookReadResult namedRanges =
        DefaultGridGrindRequestExecutor.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.NamedRangesResult(
                "ranges",
                List.of(
                    new ExcelNamedRangeSnapshot.RangeSnapshot(
                        "BudgetTotal",
                        new ExcelNamedRangeScope.WorkbookScope(),
                        "Budget!$B$4",
                        new ExcelNamedRangeTarget("Budget", "B4")))));
    WorkbookReadResult sheetSummary =
        DefaultGridGrindRequestExecutor.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.SheetSummaryResult(
                "sheet",
                new dev.erst.gridgrind.excel.WorkbookReadResult.SheetSummary("Budget", 4, 3, 2)));
    WorkbookReadResult cells =
        DefaultGridGrindRequestExecutor.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.CellsResult(
                "cells", "Budget", List.of(blank)));
    WorkbookReadResult window =
        DefaultGridGrindRequestExecutor.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.WindowResult(
                "window",
                new dev.erst.gridgrind.excel.WorkbookReadResult.Window(
                    "Budget",
                    "A1",
                    1,
                    1,
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookReadResult.WindowRow(
                            0, List.of(blank))))));
    WorkbookReadResult merged =
        DefaultGridGrindRequestExecutor.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.MergedRegionsResult(
                "merged",
                "Budget",
                List.of(new dev.erst.gridgrind.excel.WorkbookReadResult.MergedRegion("A1:B2"))));
    WorkbookReadResult hyperlinks =
        DefaultGridGrindRequestExecutor.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.HyperlinksResult(
                "hyperlinks",
                "Budget",
                List.of(
                    new dev.erst.gridgrind.excel.WorkbookReadResult.CellHyperlink(
                        "A1", new ExcelHyperlink.Url("https://example.com/report")))));
    WorkbookReadResult comments =
        DefaultGridGrindRequestExecutor.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.CommentsResult(
                "comments",
                "Budget",
                List.of(
                    new dev.erst.gridgrind.excel.WorkbookReadResult.CellComment(
                        "A1", new ExcelComment("Review", "GridGrind", false)))));
    WorkbookReadResult layout =
        DefaultGridGrindRequestExecutor.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.SheetLayoutResult(
                "layout",
                new dev.erst.gridgrind.excel.WorkbookReadResult.SheetLayout(
                    "Budget",
                    new dev.erst.gridgrind.excel.WorkbookReadResult.FreezePane.Frozen(1, 1, 1, 1),
                    List.of(new dev.erst.gridgrind.excel.WorkbookReadResult.ColumnLayout(0, 12.5)),
                    List.of(new dev.erst.gridgrind.excel.WorkbookReadResult.RowLayout(0, 18.0)))));
    WorkbookReadResult formulaSurface =
        DefaultGridGrindRequestExecutor.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.FormulaSurfaceResult(
                "formula",
                new dev.erst.gridgrind.excel.WorkbookReadResult.FormulaSurface(
                    1,
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookReadResult.SheetFormulaSurface(
                            "Budget",
                            1,
                            1,
                            List.of(
                                new dev.erst.gridgrind.excel.WorkbookReadResult.FormulaPattern(
                                    "SUM(B2:B3)", 1, List.of("B4"))))))));
    WorkbookReadResult schema =
        DefaultGridGrindRequestExecutor.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.SheetSchemaResult(
                "schema",
                new dev.erst.gridgrind.excel.WorkbookReadResult.SheetSchema(
                    "Budget",
                    "A1",
                    3,
                    2,
                    2,
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookReadResult.SchemaColumn(
                            0,
                            "A1",
                            "Item",
                            2,
                            0,
                            List.of(
                                new dev.erst.gridgrind.excel.WorkbookReadResult.TypeCount(
                                    "STRING", 2)),
                            "STRING")))));
    WorkbookReadResult namedRangeSurface =
        DefaultGridGrindRequestExecutor.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.NamedRangeSurfaceResult(
                "surface",
                new dev.erst.gridgrind.excel.WorkbookReadResult.NamedRangeSurface(
                    1,
                    0,
                    1,
                    0,
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookReadResult.NamedRangeSurfaceEntry(
                            "BudgetTotal",
                            new ExcelNamedRangeScope.WorkbookScope(),
                            "Budget!$B$4",
                            dev.erst.gridgrind.excel.WorkbookReadResult.NamedRangeBackingKind
                                .RANGE)))));

    assertInstanceOf(WorkbookReadResult.WorkbookSummaryResult.class, workbookSummary);
    assertInstanceOf(WorkbookReadResult.NamedRangesResult.class, namedRanges);
    assertInstanceOf(WorkbookReadResult.SheetSummaryResult.class, sheetSummary);
    assertInstanceOf(WorkbookReadResult.CellsResult.class, cells);
    assertInstanceOf(WorkbookReadResult.WindowResult.class, window);
    assertInstanceOf(WorkbookReadResult.MergedRegionsResult.class, merged);
    assertInstanceOf(WorkbookReadResult.HyperlinksResult.class, hyperlinks);
    assertInstanceOf(WorkbookReadResult.CommentsResult.class, comments);
    assertInstanceOf(WorkbookReadResult.SheetLayoutResult.class, layout);
    assertInstanceOf(WorkbookReadResult.FormulaSurfaceResult.class, formulaSurface);
    assertInstanceOf(WorkbookReadResult.SheetSchemaResult.class, schema);
    assertInstanceOf(WorkbookReadResult.NamedRangeSurfaceResult.class, namedRangeSurface);
    assertEquals(
        "Budget",
        cast(WorkbookReadResult.WorkbookSummaryResult.class, workbookSummary)
            .workbook()
            .sheetNames()
            .getFirst());
    assertEquals(
        "BudgetTotal",
        cast(WorkbookReadResult.NamedRangesResult.class, namedRanges)
            .namedRanges()
            .getFirst()
            .name());
    assertEquals(
        "Budget",
        cast(WorkbookReadResult.SheetSummaryResult.class, sheetSummary).sheet().sheetName());
    assertEquals(
        "A1", cast(WorkbookReadResult.CellsResult.class, cells).cells().getFirst().address());
    assertEquals(
        "A1",
        cast(WorkbookReadResult.WindowResult.class, window)
            .window()
            .rows()
            .getFirst()
            .cells()
            .getFirst()
            .address());
    assertEquals(
        "A1:B2",
        cast(WorkbookReadResult.MergedRegionsResult.class, merged)
            .mergedRegions()
            .getFirst()
            .range());
    assertEquals(
        "https://example.com/report",
        cast(WorkbookReadResult.HyperlinksResult.class, hyperlinks)
            .hyperlinks()
            .getFirst()
            .hyperlink()
            .target());
    assertEquals(
        "Review",
        cast(WorkbookReadResult.CommentsResult.class, comments)
            .comments()
            .getFirst()
            .comment()
            .text());
    assertInstanceOf(
        GridGrindResponse.FreezePaneReport.Frozen.class,
        cast(WorkbookReadResult.SheetLayoutResult.class, layout).layout().freezePanes());
    assertEquals(
        1,
        cast(WorkbookReadResult.FormulaSurfaceResult.class, formulaSurface)
            .analysis()
            .totalFormulaCellCount());
    assertEquals(
        "STRING",
        cast(WorkbookReadResult.SheetSchemaResult.class, schema)
            .analysis()
            .columns()
            .getFirst()
            .dominantType());
    assertEquals(
        GridGrindResponse.NamedRangeBackingKind.RANGE,
        cast(WorkbookReadResult.NamedRangeSurfaceResult.class, namedRangeSurface)
            .analysis()
            .namedRanges()
            .getFirst()
            .kind());
  }

  @Test
  void convertsNamedRangeFormulaSnapshotsAndFormulaBackedSurfaceEntries() {
    GridGrindResponse.NamedRangeReport formulaReport =
        DefaultGridGrindRequestExecutor.toNamedRangeReport(
            new ExcelNamedRangeSnapshot.FormulaSnapshot(
                "BudgetRollup", new ExcelNamedRangeScope.WorkbookScope(), "SUM(Budget!$B$2:$B$3)"));
    assertInstanceOf(GridGrindResponse.NamedRangeReport.FormulaReport.class, formulaReport);

    WorkbookReadResult.NamedRangeSurfaceResult surface =
        cast(
            WorkbookReadResult.NamedRangeSurfaceResult.class,
            DefaultGridGrindRequestExecutor.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookReadResult.NamedRangeSurfaceResult(
                    "surface",
                    new dev.erst.gridgrind.excel.WorkbookReadResult.NamedRangeSurface(
                        0,
                        1,
                        0,
                        1,
                        List.of(
                            new dev.erst.gridgrind.excel.WorkbookReadResult.NamedRangeSurfaceEntry(
                                "BudgetRollup",
                                new ExcelNamedRangeScope.SheetScope("Budget"),
                                "SUM(Budget!$B$2:$B$3)",
                                dev.erst.gridgrind.excel.WorkbookReadResult.NamedRangeBackingKind
                                    .FORMULA))))));

    assertEquals(
        GridGrindResponse.NamedRangeBackingKind.FORMULA,
        surface.analysis().namedRanges().getFirst().kind());
  }

  @Test
  void convertsFreezePaneNoneIntoProtocolReport() {
    WorkbookReadResult.SheetLayoutResult layout =
        cast(
            WorkbookReadResult.SheetLayoutResult.class,
            DefaultGridGrindRequestExecutor.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookReadResult.SheetLayoutResult(
                    "layout",
                    new dev.erst.gridgrind.excel.WorkbookReadResult.SheetLayout(
                        "Budget",
                        new dev.erst.gridgrind.excel.WorkbookReadResult.FreezePane.None(),
                        List.of(),
                        List.of()))));

    assertInstanceOf(GridGrindResponse.FreezePaneReport.None.class, layout.layout().freezePanes());
  }

  @Test
  void readTypeReturnsDiscriminatorsForAllReadVariants() {
    assertEquals(
        List.of(
            "GET_WORKBOOK_SUMMARY",
            "GET_NAMED_RANGES",
            "GET_SHEET_SUMMARY",
            "GET_CELLS",
            "GET_WINDOW",
            "GET_MERGED_REGIONS",
            "GET_HYPERLINKS",
            "GET_COMMENTS",
            "GET_SHEET_LAYOUT",
            "ANALYZE_FORMULA_SURFACE",
            "ANALYZE_SHEET_SCHEMA",
            "ANALYZE_NAMED_RANGE_SURFACE"),
        List.of(
            DefaultGridGrindRequestExecutor.readType(
                new WorkbookReadOperation.GetWorkbookSummary("workbook")),
            DefaultGridGrindRequestExecutor.readType(
                new WorkbookReadOperation.GetNamedRanges("ranges", new NamedRangeSelection.All())),
            DefaultGridGrindRequestExecutor.readType(
                new WorkbookReadOperation.GetSheetSummary("sheet", "Budget")),
            DefaultGridGrindRequestExecutor.readType(
                new WorkbookReadOperation.GetCells("cells", "Budget", List.of("A1"))),
            DefaultGridGrindRequestExecutor.readType(
                new WorkbookReadOperation.GetWindow("window", "Budget", "A1", 1, 1)),
            DefaultGridGrindRequestExecutor.readType(
                new WorkbookReadOperation.GetMergedRegions("merged", "Budget")),
            DefaultGridGrindRequestExecutor.readType(
                new WorkbookReadOperation.GetHyperlinks(
                    "hyperlinks", "Budget", new CellSelection.AllUsedCells())),
            DefaultGridGrindRequestExecutor.readType(
                new WorkbookReadOperation.GetComments(
                    "comments", "Budget", new CellSelection.AllUsedCells())),
            DefaultGridGrindRequestExecutor.readType(
                new WorkbookReadOperation.GetSheetLayout("layout", "Budget")),
            DefaultGridGrindRequestExecutor.readType(
                new WorkbookReadOperation.AnalyzeFormulaSurface(
                    "formula", new SheetSelection.All())),
            DefaultGridGrindRequestExecutor.readType(
                new WorkbookReadOperation.AnalyzeSheetSchema("schema", "Budget", "A1", 1, 1)),
            DefaultGridGrindRequestExecutor.readType(
                new WorkbookReadOperation.AnalyzeNamedRangeSurface(
                    "surface", new NamedRangeSelection.All()))));
  }

  @Test
  void extractsContextForReadOperations() {
    NamedRangeNotFoundException missingNamedRange =
        new NamedRangeNotFoundException("BudgetTotal", new ExcelNamedRangeScope.WorkbookScope());
    InvalidCellAddressException invalidAddress =
        new InvalidCellAddressException("BAD!", new IllegalArgumentException("bad"));
    RuntimeException runtimeException = new RuntimeException("x");

    WorkbookReadOperation workbook = new WorkbookReadOperation.GetWorkbookSummary("workbook");
    WorkbookReadOperation namedRanges =
        new WorkbookReadOperation.GetNamedRanges("ranges", new NamedRangeSelection.All());
    WorkbookReadOperation sheet = new WorkbookReadOperation.GetSheetSummary("sheet", "Budget");
    WorkbookReadOperation cells =
        new WorkbookReadOperation.GetCells("cells", "Budget", List.of("A1"));
    WorkbookReadOperation window =
        new WorkbookReadOperation.GetWindow("window", "Budget", "B2", 2, 2);
    WorkbookReadOperation merged = new WorkbookReadOperation.GetMergedRegions("merged", "Budget");
    WorkbookReadOperation hyperlinks =
        new WorkbookReadOperation.GetHyperlinks(
            "hyperlinks", "Budget", new CellSelection.AllUsedCells());
    WorkbookReadOperation comments =
        new WorkbookReadOperation.GetComments(
            "comments", "Budget", new CellSelection.Selected(List.of("A1")));
    WorkbookReadOperation layout = new WorkbookReadOperation.GetSheetLayout("layout", "Budget");
    WorkbookReadOperation formula =
        new WorkbookReadOperation.AnalyzeFormulaSurface("formula", new SheetSelection.All());
    WorkbookReadOperation schema =
        new WorkbookReadOperation.AnalyzeSheetSchema("schema", "Budget", "C3", 2, 2);
    WorkbookReadOperation surface =
        new WorkbookReadOperation.AnalyzeNamedRangeSurface(
            "surface", new NamedRangeSelection.All());

    assertNull(DefaultGridGrindRequestExecutor.sheetNameFor(workbook));
    assertNull(DefaultGridGrindRequestExecutor.addressFor(workbook, runtimeException));
    assertNull(DefaultGridGrindRequestExecutor.namedRangeNameFor(workbook, runtimeException));
    assertEquals(
        "BudgetTotal",
        DefaultGridGrindRequestExecutor.namedRangeNameFor(workbook, missingNamedRange));

    assertNull(DefaultGridGrindRequestExecutor.sheetNameFor(namedRanges));
    assertNull(DefaultGridGrindRequestExecutor.addressFor(namedRanges, runtimeException));
    assertNull(DefaultGridGrindRequestExecutor.namedRangeNameFor(namedRanges, runtimeException));
    assertEquals(
        "BudgetTotal",
        DefaultGridGrindRequestExecutor.namedRangeNameFor(namedRanges, missingNamedRange));

    assertEquals("Budget", DefaultGridGrindRequestExecutor.sheetNameFor(sheet));
    assertNull(DefaultGridGrindRequestExecutor.addressFor(sheet, runtimeException));
    assertNull(DefaultGridGrindRequestExecutor.namedRangeNameFor(sheet, runtimeException));

    assertEquals("Budget", DefaultGridGrindRequestExecutor.sheetNameFor(cells));
    assertEquals("BAD!", DefaultGridGrindRequestExecutor.addressFor(cells, invalidAddress));
    assertNull(DefaultGridGrindRequestExecutor.addressFor(cells, runtimeException));
    assertNull(DefaultGridGrindRequestExecutor.namedRangeNameFor(cells, runtimeException));

    assertEquals("Budget", DefaultGridGrindRequestExecutor.sheetNameFor(window));
    assertEquals("B2", DefaultGridGrindRequestExecutor.addressFor(window, runtimeException));
    assertNull(DefaultGridGrindRequestExecutor.namedRangeNameFor(window, runtimeException));

    assertEquals("Budget", DefaultGridGrindRequestExecutor.sheetNameFor(merged));
    assertNull(DefaultGridGrindRequestExecutor.addressFor(merged, runtimeException));
    assertNull(DefaultGridGrindRequestExecutor.namedRangeNameFor(merged, runtimeException));

    assertEquals("Budget", DefaultGridGrindRequestExecutor.sheetNameFor(hyperlinks));
    assertNull(DefaultGridGrindRequestExecutor.addressFor(hyperlinks, runtimeException));
    assertNull(DefaultGridGrindRequestExecutor.namedRangeNameFor(hyperlinks, runtimeException));

    assertEquals("Budget", DefaultGridGrindRequestExecutor.sheetNameFor(comments));
    assertNull(DefaultGridGrindRequestExecutor.addressFor(comments, runtimeException));
    assertNull(DefaultGridGrindRequestExecutor.namedRangeNameFor(comments, runtimeException));

    assertEquals("Budget", DefaultGridGrindRequestExecutor.sheetNameFor(layout));
    assertNull(DefaultGridGrindRequestExecutor.addressFor(layout, runtimeException));
    assertNull(DefaultGridGrindRequestExecutor.namedRangeNameFor(layout, runtimeException));

    assertNull(DefaultGridGrindRequestExecutor.sheetNameFor(formula));
    assertNull(DefaultGridGrindRequestExecutor.addressFor(formula, runtimeException));
    assertNull(DefaultGridGrindRequestExecutor.namedRangeNameFor(formula, runtimeException));

    assertEquals("Budget", DefaultGridGrindRequestExecutor.sheetNameFor(schema));
    assertEquals("C3", DefaultGridGrindRequestExecutor.addressFor(schema, runtimeException));
    assertNull(DefaultGridGrindRequestExecutor.namedRangeNameFor(schema, runtimeException));

    assertNull(DefaultGridGrindRequestExecutor.sheetNameFor(surface));
    assertNull(DefaultGridGrindRequestExecutor.addressFor(surface, runtimeException));
    assertNull(DefaultGridGrindRequestExecutor.namedRangeNameFor(surface, runtimeException));
  }

  @Test
  void extractsReadContextFromExceptionsBeforeFallingBackToReadShape() {
    InvalidCellAddressException invalidAddress =
        new InvalidCellAddressException("BAD!", new IllegalArgumentException("bad"));
    NamedRangeNotFoundException missingNamedRange =
        new NamedRangeNotFoundException("BudgetTotal", new ExcelNamedRangeScope.WorkbookScope());

    assertEquals(
        "BAD!",
        DefaultGridGrindRequestExecutor.addressFor(
            new WorkbookReadOperation.AnalyzeSheetSchema("schema", "Budget", "C3", 2, 2),
            invalidAddress));
    assertEquals(
        "BudgetTotal",
        DefaultGridGrindRequestExecutor.namedRangeNameFor(
            new WorkbookReadOperation.AnalyzeNamedRangeSurface(
                "surface", new NamedRangeSelection.All()),
            missingNamedRange));
  }

  @Test
  void extractsContextForAuthoringMetadataAndNamedRangeOperations() {
    RuntimeException exception = new RuntimeException("test");
    WorkbookOperation setHyperlink =
        new WorkbookOperation.SetHyperlink(
            "Budget", "A1", new HyperlinkTarget.Url("https://example.com/report"));
    WorkbookOperation clearHyperlink = new WorkbookOperation.ClearHyperlink("Budget", "A1");
    WorkbookOperation setComment =
        new WorkbookOperation.SetComment(
            "Budget", "A1", new CommentInput("Review", "GridGrind", false));
    WorkbookOperation clearComment = new WorkbookOperation.ClearComment("Budget", "A1");
    WorkbookOperation applyStyle =
        new WorkbookOperation.ApplyStyle(
            "Budget",
            "A1:B2",
            new CellStyleInput(
                null, true, null, null, null, null, null, null, null, null, null, null, null));
    WorkbookOperation setNamedRange =
        new WorkbookOperation.SetNamedRange(
            "BudgetTotal", new NamedRangeScope.Workbook(), new NamedRangeTarget("Budget", "B4"));
    WorkbookOperation deleteNamedRangeWorkbook =
        new WorkbookOperation.DeleteNamedRange("BudgetTotal", new NamedRangeScope.Workbook());
    WorkbookOperation deleteNamedRangeSheet =
        new WorkbookOperation.DeleteNamedRange("LocalItem", new NamedRangeScope.Sheet("Budget"));

    assertNull(DefaultGridGrindRequestExecutor.formulaFor(setHyperlink, exception));
    assertNull(DefaultGridGrindRequestExecutor.formulaFor(clearHyperlink, exception));
    assertNull(DefaultGridGrindRequestExecutor.formulaFor(setComment, exception));
    assertNull(DefaultGridGrindRequestExecutor.formulaFor(clearComment, exception));
    assertNull(DefaultGridGrindRequestExecutor.formulaFor(applyStyle, exception));
    assertNull(DefaultGridGrindRequestExecutor.formulaFor(setNamedRange, exception));
    assertNull(DefaultGridGrindRequestExecutor.formulaFor(deleteNamedRangeWorkbook, exception));

    assertEquals("Budget", DefaultGridGrindRequestExecutor.sheetNameFor(setHyperlink, exception));
    assertEquals("Budget", DefaultGridGrindRequestExecutor.sheetNameFor(clearHyperlink, exception));
    assertEquals("Budget", DefaultGridGrindRequestExecutor.sheetNameFor(setComment, exception));
    assertEquals("Budget", DefaultGridGrindRequestExecutor.sheetNameFor(clearComment, exception));
    assertEquals("Budget", DefaultGridGrindRequestExecutor.sheetNameFor(applyStyle, exception));
    assertEquals("Budget", DefaultGridGrindRequestExecutor.sheetNameFor(setNamedRange, exception));
    assertNull(DefaultGridGrindRequestExecutor.sheetNameFor(deleteNamedRangeWorkbook, exception));
    assertEquals(
        "Budget", DefaultGridGrindRequestExecutor.sheetNameFor(deleteNamedRangeSheet, exception));

    assertEquals("A1", DefaultGridGrindRequestExecutor.addressFor(setHyperlink, exception));
    assertEquals("A1", DefaultGridGrindRequestExecutor.addressFor(clearHyperlink, exception));
    assertEquals("A1", DefaultGridGrindRequestExecutor.addressFor(setComment, exception));
    assertEquals("A1", DefaultGridGrindRequestExecutor.addressFor(clearComment, exception));
    assertNull(DefaultGridGrindRequestExecutor.addressFor(setNamedRange, exception));

    assertEquals("A1:B2", DefaultGridGrindRequestExecutor.rangeFor(applyStyle, exception));
    assertEquals("B4", DefaultGridGrindRequestExecutor.rangeFor(setNamedRange, exception));
    assertNull(DefaultGridGrindRequestExecutor.rangeFor(deleteNamedRangeWorkbook, exception));

    assertEquals(
        "BudgetTotal", DefaultGridGrindRequestExecutor.namedRangeNameFor(setNamedRange, exception));
    assertEquals(
        "BudgetTotal",
        DefaultGridGrindRequestExecutor.namedRangeNameFor(deleteNamedRangeWorkbook, exception));
    assertEquals(
        "LocalItem",
        DefaultGridGrindRequestExecutor.namedRangeNameFor(deleteNamedRangeSheet, exception));
    assertEquals(
        "BudgetTotal",
        DefaultGridGrindRequestExecutor.namedRangeNameFor(
            new WorkbookOperation.EnsureSheet("Budget"),
            new NamedRangeNotFoundException(
                "BudgetTotal", new ExcelNamedRangeScope.WorkbookScope())));
  }

  @Test
  void extractsAddressAndRangeFallbacksAndExhaustsNamedRangeNullArms() {
    InvalidFormulaException invalidFormula =
        new InvalidFormulaException(
            "Budget", "C3", "SUM(", "bad formula", new IllegalArgumentException("bad"));
    NamedRangeNotFoundException missingNamedRange =
        new NamedRangeNotFoundException("BudgetTotal", new ExcelNamedRangeScope.WorkbookScope());
    InvalidCellAddressException invalidAddress =
        new InvalidCellAddressException("BAD!", new IllegalArgumentException("bad"));
    InvalidRangeAddressException invalidRange =
        new InvalidRangeAddressException("A1:", new IllegalArgumentException("bad"));

    assertEquals(
        "C3",
        DefaultGridGrindRequestExecutor.addressFor(
            new WorkbookOperation.EnsureSheet("Budget"), invalidFormula));
    assertEquals(
        "BAD!",
        DefaultGridGrindRequestExecutor.addressFor(
            new WorkbookOperation.DeleteNamedRange("BudgetTotal", new NamedRangeScope.Workbook()),
            invalidAddress));

    assertEquals(
        "C1:D2",
        DefaultGridGrindRequestExecutor.rangeFor(
            new WorkbookOperation.SetRange(
                "Budget", "C1:D2", List.of(List.of(new CellInput.Text("x")))),
            invalidFormula));
    assertEquals(
        "E1:E2",
        DefaultGridGrindRequestExecutor.rangeFor(
            new WorkbookOperation.ClearRange("Budget", "E1:E2"), invalidFormula));
    assertNull(
        DefaultGridGrindRequestExecutor.rangeFor(
            new WorkbookOperation.SetHyperlink(
                "Budget", "A1", new HyperlinkTarget.Url("https://example.com/report")),
            invalidFormula));
    assertNull(
        DefaultGridGrindRequestExecutor.rangeFor(
            new WorkbookOperation.ClearHyperlink("Budget", "A1"), invalidFormula));
    assertNull(
        DefaultGridGrindRequestExecutor.rangeFor(
            new WorkbookOperation.SetComment(
                "Budget", "A1", new CommentInput("Review", "GridGrind", false)),
            invalidFormula));
    assertNull(
        DefaultGridGrindRequestExecutor.rangeFor(
            new WorkbookOperation.ClearComment("Budget", "A1"), invalidFormula));
    assertEquals(
        "A1:",
        DefaultGridGrindRequestExecutor.rangeFor(
            new WorkbookOperation.EnsureSheet("Budget"), invalidRange));

    List<WorkbookOperation> operationsWithoutNamedRanges =
        List.of(
            new WorkbookOperation.EnsureSheet("Budget"),
            new WorkbookOperation.RenameSheet("Budget", "Summary"),
            new WorkbookOperation.DeleteSheet("Budget"),
            new WorkbookOperation.MoveSheet("Budget", 0),
            new WorkbookOperation.MergeCells("Budget", "A1:B2"),
            new WorkbookOperation.UnmergeCells("Budget", "A1:B2"),
            new WorkbookOperation.SetColumnWidth("Budget", 0, 1, 16.0),
            new WorkbookOperation.SetRowHeight("Budget", 0, 1, 28.5),
            new WorkbookOperation.FreezePanes("Budget", 1, 1, 1, 1),
            new WorkbookOperation.SetCell("Budget", "A1", new CellInput.Text("x")),
            new WorkbookOperation.SetRange(
                "Budget", "A1:B1", List.of(List.of(new CellInput.Text("x")))),
            new WorkbookOperation.ClearRange("Budget", "A1:B1"),
            new WorkbookOperation.SetHyperlink(
                "Budget", "A1", new HyperlinkTarget.Url("https://example.com/report")),
            new WorkbookOperation.ClearHyperlink("Budget", "A1"),
            new WorkbookOperation.SetComment(
                "Budget", "A1", new CommentInput("Review", "GridGrind", false)),
            new WorkbookOperation.ClearComment("Budget", "A1"),
            new WorkbookOperation.ApplyStyle(
                "Budget",
                "A1:B2",
                new CellStyleInput(
                    null, true, null, null, null, null, null, null, null, null, null, null, null)),
            new WorkbookOperation.AppendRow("Budget", List.of(new CellInput.Text("x"))),
            new WorkbookOperation.AutoSizeColumns("Budget"),
            new WorkbookOperation.EvaluateFormulas(),
            new WorkbookOperation.ForceFormulaRecalculationOnOpen());

    for (WorkbookOperation operation : operationsWithoutNamedRanges) {
      assertNull(DefaultGridGrindRequestExecutor.namedRangeNameFor(operation, invalidFormula));
    }

    assertEquals(
        "BudgetTotal",
        DefaultGridGrindRequestExecutor.namedRangeNameFor(
            new WorkbookOperation.FreezePanes("Budget", 1, 1, 1, 1), missingNamedRange));
  }

  private static GridGrindRequest request(
      GridGrindRequest.WorkbookSource source,
      GridGrindRequest.WorkbookPersistence persistence,
      List<WorkbookOperation> operations,
      WorkbookReadOperation... reads) {
    return new GridGrindRequest(source, persistence, operations, List.of(reads));
  }

  private static GridGrindResponse.Success success(GridGrindResponse response) {
    return cast(GridGrindResponse.Success.class, response);
  }

  private static GridGrindResponse.Failure failure(GridGrindResponse response) {
    return cast(GridGrindResponse.Failure.class, response);
  }

  private static String savedPath(GridGrindResponse.Success success) {
    return cast(GridGrindResponse.PersistenceOutcome.Saved.class, success.persistence()).path();
  }

  private static List<String> requestIds(GridGrindResponse.Success success) {
    return success.reads().stream().map(WorkbookReadResult::requestId).toList();
  }

  private static <T> T cast(Class<T> type, Object value) {
    return type.cast(assertInstanceOf(type, value));
  }

  private static <T extends WorkbookReadResult> T read(
      GridGrindResponse.Success success, String requestId, Class<T> type) {
    return cast(
        type,
        success.reads().stream()
            .filter(result -> result.requestId().equals(requestId))
            .findFirst()
            .orElseThrow(
                () -> new AssertionError("Missing read result for requestId " + requestId)));
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

  private static ExcelCellStyleSnapshot defaultStyle() {
    return new ExcelCellStyleSnapshot(
        "",
        false,
        false,
        false,
        ExcelHorizontalAlignment.GENERAL,
        ExcelVerticalAlignment.BOTTOM,
        "Aptos",
        ExcelFontHeight.fromPoints(new BigDecimal("11")),
        null,
        false,
        false,
        null,
        ExcelBorderStyle.NONE,
        ExcelBorderStyle.NONE,
        ExcelBorderStyle.NONE,
        ExcelBorderStyle.NONE);
  }
}
