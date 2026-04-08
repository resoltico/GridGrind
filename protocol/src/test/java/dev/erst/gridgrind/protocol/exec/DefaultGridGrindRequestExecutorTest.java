package dev.erst.gridgrind.protocol.exec;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.ExcelAutofilterSnapshot;
import dev.erst.gridgrind.excel.ExcelBorderSide;
import dev.erst.gridgrind.excel.ExcelBorderSnapshot;
import dev.erst.gridgrind.excel.ExcelBorderStyle;
import dev.erst.gridgrind.excel.ExcelCellAlignmentSnapshot;
import dev.erst.gridgrind.excel.ExcelCellFillSnapshot;
import dev.erst.gridgrind.excel.ExcelCellFontSnapshot;
import dev.erst.gridgrind.excel.ExcelCellMetadataSnapshot;
import dev.erst.gridgrind.excel.ExcelCellProtectionSnapshot;
import dev.erst.gridgrind.excel.ExcelCellSnapshot;
import dev.erst.gridgrind.excel.ExcelCellStyleSnapshot;
import dev.erst.gridgrind.excel.ExcelCellValue;
import dev.erst.gridgrind.excel.ExcelComment;
import dev.erst.gridgrind.excel.ExcelComparisonOperator;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingBlockSnapshot;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingRuleSnapshot;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingThresholdSnapshot;
import dev.erst.gridgrind.excel.ExcelConditionalFormattingThresholdType;
import dev.erst.gridgrind.excel.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.excel.ExcelDataValidationSnapshot;
import dev.erst.gridgrind.excel.ExcelDifferentialStyleSnapshot;
import dev.erst.gridgrind.excel.ExcelFontHeight;
import dev.erst.gridgrind.excel.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.ExcelHyperlink;
import dev.erst.gridgrind.excel.ExcelNamedRangeDefinition;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import dev.erst.gridgrind.excel.ExcelNamedRangeSnapshot;
import dev.erst.gridgrind.excel.ExcelNamedRangeTarget;
import dev.erst.gridgrind.excel.ExcelPaneRegion;
import dev.erst.gridgrind.excel.ExcelPrintOrientation;
import dev.erst.gridgrind.excel.ExcelSheetCopyPosition;
import dev.erst.gridgrind.excel.ExcelSheetPane;
import dev.erst.gridgrind.excel.ExcelSheetProtectionSettings;
import dev.erst.gridgrind.excel.ExcelSheetVisibility;
import dev.erst.gridgrind.excel.ExcelTableSelection;
import dev.erst.gridgrind.excel.ExcelTableSnapshot;
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
import dev.erst.gridgrind.excel.WorkbookLocation;
import dev.erst.gridgrind.excel.WorkbookNotFoundException;
import dev.erst.gridgrind.excel.WorkbookReadCommand;
import dev.erst.gridgrind.excel.WorkbookReadExecutor;
import dev.erst.gridgrind.excel.XlsxRoundTrip;
import dev.erst.gridgrind.protocol.dto.*;
import dev.erst.gridgrind.protocol.operation.WorkbookOperation;
import dev.erst.gridgrind.protocol.read.*;
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
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidation;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STDataValidationType;

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
    assertEquals(List.of(), success.warnings());
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
  void executesRichTextCellWorkflowAndReportsStructuredRuns() {
    GridGrindResponse response =
        new DefaultGridGrindRequestExecutor()
            .execute(
                request(
                    new GridGrindRequest.WorkbookSource.New(),
                    new GridGrindRequest.WorkbookPersistence.None(),
                    List.of(
                        new WorkbookOperation.EnsureSheet("Budget"),
                        new WorkbookOperation.ApplyStyle(
                            "Budget",
                            "A1",
                            new CellStyleInput(
                                null,
                                null,
                                new CellFontInput(
                                    null,
                                    Boolean.TRUE,
                                    "Aptos",
                                    new FontHeightInput.Twips(260),
                                    "#112233",
                                    null,
                                    null),
                                null,
                                null,
                                null)),
                        new WorkbookOperation.SetCell(
                            "Budget",
                            "A1",
                            new CellInput.RichText(
                                List.of(
                                    new RichTextRunInput("Budget", null),
                                    new RichTextRunInput(
                                        " FY26",
                                        new CellFontInput(
                                            Boolean.TRUE,
                                            null,
                                            null,
                                            null,
                                            "#FF0000",
                                            null,
                                            null)))))),
                    new WorkbookReadOperation.GetCells("cells", "Budget", List.of("A1"))));

    GridGrindResponse.Success success = success(response);
    WorkbookReadResult.CellsResult cells =
        read(success, "cells", WorkbookReadResult.CellsResult.class);
    GridGrindResponse.CellReport.TextReport cell =
        cast(GridGrindResponse.CellReport.TextReport.class, cells.cells().getFirst());

    assertEquals("Budget FY26", cell.stringValue());
    assertNotNull(cell.richText());
    assertEquals(2, cell.richText().size());
    assertEquals("Budget", cell.richText().get(0).text());
    assertEquals("Aptos", cell.richText().get(0).font().fontName());
    assertEquals("#112233", cell.richText().get(0).font().fontColor());
    assertTrue(cell.richText().get(0).font().italic());
    assertFalse(cell.richText().get(0).font().bold());
    assertEquals(" FY26", cell.richText().get(1).text());
    assertEquals("Aptos", cell.richText().get(1).font().fontName());
    assertEquals("#FF0000", cell.richText().get(1).font().fontColor());
    assertTrue(cell.richText().get(1).font().bold());
    assertTrue(cell.richText().get(1).font().italic());
  }

  @Test
  void surfacesRequestWarningsAlongsideSuccessfulExecution() {
    GridGrindResponse response =
        new DefaultGridGrindRequestExecutor()
            .execute(
                request(
                    new GridGrindRequest.WorkbookSource.New(),
                    new GridGrindRequest.WorkbookPersistence.None(),
                    List.of(
                        new WorkbookOperation.EnsureSheet("Budget Review"),
                        new WorkbookOperation.EnsureSheet("Summary"),
                        new WorkbookOperation.SetCell(
                            "Budget Review", "A1", new CellInput.Numeric(1200.0)),
                        new WorkbookOperation.SetCell(
                            "Summary", "A1", new CellInput.Formula("Budget Review!A1")))));

    GridGrindResponse.Success success = success(response);

    assertEquals(
        List.of(
            new RequestWarning(
                3,
                "SET_CELL",
                "Formula references same-request sheet names with spaces without single quotes: Budget Review. Use 'Sheet Name'!A1 syntax.")),
        success.warnings());
    assertEquals(List.of(), success.reads());
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
                        new WorkbookOperation.SetSheetPane(
                            "Budget", new PaneInput.Frozen(1, 1, 1, 1)),
                        new WorkbookOperation.SetSheetZoom("Budget", 125),
                        new WorkbookOperation.SetPrintLayout(
                            "Budget",
                            new PrintLayoutInput(
                                new PrintAreaInput.Range("A1:B12"),
                                ExcelPrintOrientation.LANDSCAPE,
                                new PrintScalingInput.Fit(1, 0),
                                new PrintTitleRowsInput.Band(0, 0),
                                new PrintTitleColumnsInput.Band(0, 0),
                                new HeaderFooterTextInput("Budget", "", ""),
                                new HeaderFooterTextInput("", "Page &P", "")))),
                    new WorkbookReadOperation.GetCells("cells", "Budget", List.of("A1")),
                    new WorkbookReadOperation.GetMergedRegions("merged", "Budget"),
                    new WorkbookReadOperation.GetSheetLayout("layout", "Budget"),
                    new WorkbookReadOperation.GetPrintLayout("printLayout", "Budget"),
                    new WorkbookReadOperation.GetWorkbookSummary("workbook")));

    GridGrindResponse.Success success = success(response);
    WorkbookReadResult.CellsResult cells =
        read(success, "cells", WorkbookReadResult.CellsResult.class);
    WorkbookReadResult.MergedRegionsResult merged =
        read(success, "merged", WorkbookReadResult.MergedRegionsResult.class);
    GridGrindResponse.SheetLayoutReport layout =
        read(success, "layout", WorkbookReadResult.SheetLayoutResult.class).layout();
    PrintLayoutReport printLayout =
        read(success, "printLayout", WorkbookReadResult.PrintLayoutResult.class).layout();
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
    assertInstanceOf(PaneReport.Frozen.class, layout.pane());
    PaneReport.Frozen frozen = cast(PaneReport.Frozen.class, layout.pane());
    assertEquals(1, frozen.splitColumn());
    assertEquals(1, frozen.splitRow());
    assertEquals(125, layout.zoomPercent());
    assertEquals(16.0, layout.columns().getFirst().widthCharacters());
    assertEquals(28.5, layout.rows().getFirst().heightPoints());
    assertEquals(ExcelPrintOrientation.LANDSCAPE, printLayout.orientation());
    assertEquals(List.of("Budget"), workbook.sheetNames());

    assertEquals(List.of("A1:B1"), XlsxRoundTrip.mergedRegions(workbookPath, "Budget"));
    assertEquals(4096, XlsxRoundTrip.columnWidth(workbookPath, "Budget", 0));
    assertEquals((short) 570, XlsxRoundTrip.rowHeightTwips(workbookPath, "Budget", 0));
    assertEquals(new ExcelSheetPane.Frozen(1, 1, 1, 1), XlsxRoundTrip.pane(workbookPath, "Budget"));
    assertEquals(125, XlsxRoundTrip.zoomPercent(workbookPath, "Budget"));
  }

  @Test
  void returnsB3SheetLayoutFactsAndPersistsVisibilityGrouping() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-b3-layout-");

    GridGrindResponse response =
        new DefaultGridGrindRequestExecutor()
            .execute(
                request(
                    new GridGrindRequest.WorkbookSource.New(),
                    new GridGrindRequest.WorkbookPersistence.SaveAs(workbookPath.toString()),
                    List.of(
                        new WorkbookOperation.EnsureSheet("Layout"),
                        new WorkbookOperation.SetRange(
                            "Layout",
                            "A1:F6",
                            List.of(
                                List.of(
                                    new CellInput.Text("Item"),
                                    new CellInput.Text("Qty"),
                                    new CellInput.Text("Status"),
                                    new CellInput.Text("Note"),
                                    new CellInput.Text("Owner"),
                                    new CellInput.Text("Flag")),
                                List.of(
                                    new CellInput.Text("Hosting"),
                                    new CellInput.Numeric(42.0),
                                    new CellInput.Text("Open"),
                                    new CellInput.Text("Alpha"),
                                    new CellInput.Text("Ada"),
                                    new CellInput.Text("Y")),
                                List.of(
                                    new CellInput.Text("Support"),
                                    new CellInput.Numeric(84.0),
                                    new CellInput.Text("Closed"),
                                    new CellInput.Text("Beta"),
                                    new CellInput.Text("Lin"),
                                    new CellInput.Text("N")),
                                List.of(
                                    new CellInput.Text("Ops"),
                                    new CellInput.Numeric(168.0),
                                    new CellInput.Text("Open"),
                                    new CellInput.Text("Gamma"),
                                    new CellInput.Text("Bea"),
                                    new CellInput.Text("Y")),
                                List.of(
                                    new CellInput.Text("QA"),
                                    new CellInput.Numeric(21.0),
                                    new CellInput.Text("Queued"),
                                    new CellInput.Text("Delta"),
                                    new CellInput.Text("Kai"),
                                    new CellInput.Text("N")),
                                List.of(
                                    new CellInput.Text("Infra"),
                                    new CellInput.Numeric(7.0),
                                    new CellInput.Text("Done"),
                                    new CellInput.Text("Epsilon"),
                                    new CellInput.Text("Mia"),
                                    new CellInput.Text("Y")))),
                        new WorkbookOperation.GroupRows(
                            "Layout", new RowSpanInput.Band(1, 3), true),
                        new WorkbookOperation.SetRowVisibility(
                            "Layout", new RowSpanInput.Band(5, 5), true),
                        new WorkbookOperation.GroupColumns(
                            "Layout", new ColumnSpanInput.Band(1, 3), true),
                        new WorkbookOperation.SetColumnVisibility(
                            "Layout", new ColumnSpanInput.Band(5, 5), true)),
                    new WorkbookReadOperation.GetSheetLayout("layout", "Layout")));

    GridGrindResponse.Success success = success(response);
    GridGrindResponse.SheetLayoutReport layout =
        read(success, "layout", WorkbookReadResult.SheetLayoutResult.class).layout();

    assertEquals(workbookPath.toAbsolutePath().toString(), savedPath(success));
    assertEquals(6, layout.rows().size());
    assertTrue(layout.rows().get(1).hidden());
    assertEquals(1, layout.rows().get(1).outlineLevel());
    assertTrue(layout.rows().get(4).collapsed());
    assertTrue(layout.rows().get(5).hidden());
    assertEquals(6, layout.columns().size());
    assertTrue(layout.columns().get(1).hidden());
    assertEquals(1, layout.columns().get(1).outlineLevel());
    assertTrue(layout.columns().get(4).collapsed());
    assertTrue(layout.columns().get(5).hidden());

    dev.erst.gridgrind.excel.WorkbookReadResult.SheetLayout reopenedLayout =
        XlsxRoundTrip.sheetLayout(workbookPath, "Layout");
    assertTrue(reopenedLayout.rows().get(1).hidden());
    assertEquals(1, reopenedLayout.rows().get(1).outlineLevel());
    assertTrue(reopenedLayout.rows().get(4).collapsed());
    assertTrue(reopenedLayout.rows().get(5).hidden());
    assertTrue(reopenedLayout.columns().get(1).hidden());
    assertEquals(1, reopenedLayout.columns().get(1).outlineLevel());
    assertTrue(reopenedLayout.columns().get(4).collapsed());
    assertTrue(reopenedLayout.columns().get(5).hidden());
  }

  @Test
  void executesB3InsertDeleteAndShiftOperationsAndPersistsMovedCells() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-b3-geometry-");

    GridGrindResponse response =
        new DefaultGridGrindRequestExecutor()
            .execute(
                request(
                    new GridGrindRequest.WorkbookSource.New(),
                    new GridGrindRequest.WorkbookPersistence.SaveAs(workbookPath.toString()),
                    List.of(
                        new WorkbookOperation.EnsureSheet("Moves"),
                        new WorkbookOperation.SetRange(
                            "Moves",
                            "A1:D3",
                            List.of(
                                List.of(
                                    new CellInput.Text("Item"),
                                    new CellInput.Text("Qty"),
                                    new CellInput.Text("Status"),
                                    new CellInput.Text("Note")),
                                List.of(
                                    new CellInput.Text("Hosting"),
                                    new CellInput.Numeric(42.0),
                                    new CellInput.Text("Open"),
                                    new CellInput.Text("Alpha")),
                                List.of(
                                    new CellInput.Text("Support"),
                                    new CellInput.Numeric(84.0),
                                    new CellInput.Text("Closed"),
                                    new CellInput.Text("Beta")))),
                        new WorkbookOperation.InsertRows("Moves", 1, 1),
                        new WorkbookOperation.SetCell("Moves", "A2", new CellInput.Text("Spacer")),
                        new WorkbookOperation.ShiftRows("Moves", new RowSpanInput.Band(2, 3), 1),
                        new WorkbookOperation.DeleteRows("Moves", new RowSpanInput.Band(2, 2)),
                        new WorkbookOperation.InsertColumns("Moves", 1, 1),
                        new WorkbookOperation.SetCell("Moves", "B1", new CellInput.Text("Pad")),
                        new WorkbookOperation.ShiftColumns(
                            "Moves", new ColumnSpanInput.Band(2, 4), 1),
                        new WorkbookOperation.DeleteColumns(
                            "Moves", new ColumnSpanInput.Band(2, 2))),
                    new WorkbookReadOperation.GetCells(
                        "cells", "Moves", List.of("A2", "B1", "A3", "C3", "E4"))));

    GridGrindResponse.Success success = success(response);
    WorkbookReadResult.CellsResult cells =
        read(success, "cells", WorkbookReadResult.CellsResult.class);

    assertEquals(workbookPath.toAbsolutePath().toString(), savedPath(success));
    assertEquals(
        "Spacer",
        cast(GridGrindResponse.CellReport.TextReport.class, cells.cells().get(0)).stringValue());
    assertEquals(
        "Pad",
        cast(GridGrindResponse.CellReport.TextReport.class, cells.cells().get(1)).stringValue());
    assertEquals(
        "Hosting",
        cast(GridGrindResponse.CellReport.TextReport.class, cells.cells().get(2)).stringValue());
    assertEquals(
        42.0,
        cast(GridGrindResponse.CellReport.NumberReport.class, cells.cells().get(3)).numberValue());
    assertEquals(
        "Beta",
        cast(GridGrindResponse.CellReport.TextReport.class, cells.cells().get(4)).stringValue());

    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      assertEquals("Spacer", workbook.sheet("Moves").text("A2"));
      assertEquals("Pad", workbook.sheet("Moves").text("B1"));
      assertEquals("Hosting", workbook.sheet("Moves").text("A3"));
      assertEquals(42.0, workbook.sheet("Moves").number("C3"));
      assertEquals("Beta", workbook.sheet("Moves").text("E4"));
    }
  }

  @Test
  void returnsStructuredFailureForRowStructuralEditsThatWouldMoveTables() {
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
                                "Budget",
                                "A1:B3",
                                List.of(
                                    List.of(new CellInput.Text("Item"), new CellInput.Text("Qty")),
                                    List.of(
                                        new CellInput.Text("Hosting"), new CellInput.Numeric(42.0)),
                                    List.of(
                                        new CellInput.Text("Support"),
                                        new CellInput.Numeric(84.0)))),
                            new WorkbookOperation.SetTable(
                                new TableInput(
                                    "BudgetTable",
                                    "Budget",
                                    "A1:B3",
                                    false,
                                    new TableStyleInput.None())),
                            new WorkbookOperation.InsertRows("Budget", 1, 1)))));

    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertTrue(failure.problem().message().contains("table 'BudgetTable'"));
    assertTrue(failure.problem().message().contains("row structural edits"));
  }

  @Test
  void returnsStructuredFailureForColumnStructuralEditsWhenWorkbookHasFormulas() {
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
                                "Budget", "A1", new CellInput.Text("Item")),
                            new WorkbookOperation.SetCell(
                                "Budget", "B2", new CellInput.Formula("SUM(1, 1)")),
                            new WorkbookOperation.InsertColumns("Budget", 1, 1)))));

    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertTrue(failure.problem().message().contains("workbook formulas are present"));
    assertTrue(failure.problem().message().contains("column structural edits"));
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

    assertEquals(new HyperlinkTarget.Url("https://example.com/report"), linkedCell.hyperlink());
    assertEquals("Review", linkedCell.comment().text());
    assertEquals("A1", hyperlinks.hyperlinks().getFirst().address());
    assertEquals(
        new HyperlinkTarget.Url("https://example.com/report"),
        hyperlinks.hyperlinks().getFirst().hyperlink());
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
  void returnsDataValidationReadResultsAndPersistsNormalizedRules() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-data-validation-ops-");

    GridGrindResponse response =
        new DefaultGridGrindRequestExecutor()
            .execute(
                request(
                    new GridGrindRequest.WorkbookSource.New(),
                    new GridGrindRequest.WorkbookPersistence.SaveAs(workbookPath.toString()),
                    List.of(
                        new WorkbookOperation.EnsureSheet("Budget"),
                        new WorkbookOperation.SetDataValidation(
                            "Budget",
                            "A2:C5",
                            new DataValidationInput(
                                new DataValidationRuleInput.ExplicitList(List.of("Queued", "Done")),
                                true,
                                false,
                                new DataValidationPromptInput(
                                    "Status", "Pick one workflow state.", true),
                                new DataValidationErrorAlertInput(
                                    ExcelDataValidationErrorStyle.STOP,
                                    "Invalid status",
                                    "Use one of the allowed values.",
                                    true))),
                        new WorkbookOperation.ClearDataValidations(
                            "Budget", new RangeSelection.Selected(List.of("B3"))),
                        new WorkbookOperation.SetDataValidation(
                            "Budget",
                            "E2:E5",
                            new DataValidationInput(
                                new DataValidationRuleInput.FormulaList("#REF!"),
                                false,
                                false,
                                null,
                                null))),
                    new WorkbookReadOperation.GetDataValidations(
                        "validations", "Budget", new RangeSelection.All()),
                    new WorkbookReadOperation.AnalyzeDataValidationHealth(
                        "health", new SheetSelection.Selected(List.of("Budget")))));

    GridGrindResponse.Success success = success(response);
    WorkbookReadResult.DataValidationsResult validations =
        read(success, "validations", WorkbookReadResult.DataValidationsResult.class);
    WorkbookReadResult.DataValidationHealthResult health =
        read(success, "health", WorkbookReadResult.DataValidationHealthResult.class);
    List<ExcelDataValidationSnapshot> persisted =
        XlsxRoundTrip.dataValidations(workbookPath, "Budget");

    assertEquals(workbookPath.toAbsolutePath().toString(), savedPath(success));
    assertEquals(2, validations.validations().size());
    DataValidationEntryReport.Supported explicitList =
        assertInstanceOf(
            DataValidationEntryReport.Supported.class, validations.validations().getFirst());
    DataValidationEntryReport.Supported brokenFormula =
        assertInstanceOf(
            DataValidationEntryReport.Supported.class, validations.validations().get(1));
    assertEquals(List.of("A2:C2", "A4:C5", "A3", "C3"), explicitList.ranges());
    assertInstanceOf(DataValidationRuleInput.ExplicitList.class, explicitList.validation().rule());
    assertEquals(List.of("E2:E5"), brokenFormula.ranges());
    assertEquals(
        "#REF!",
        cast(DataValidationRuleInput.FormulaList.class, brokenFormula.validation().rule())
            .formula());
    assertEquals(2, health.analysis().checkedValidationCount());
    assertEquals(1, health.analysis().summary().errorCount());
    assertEquals(
        AnalysisFindingCode.DATA_VALIDATION_BROKEN_FORMULA,
        health.analysis().findings().getFirst().code());
    assertEquals(2, persisted.size());
    assertEquals(List.of("A2:C2", "A4:C5", "A3", "C3"), persisted.getFirst().ranges());
  }

  @Test
  void returnsConditionalFormattingReadResultsAndPersistsDefinitions() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-conditional-formatting-");

    GridGrindResponse response =
        new DefaultGridGrindRequestExecutor()
            .execute(
                request(
                    new GridGrindRequest.WorkbookSource.New(),
                    new GridGrindRequest.WorkbookPersistence.SaveAs(workbookPath.toString()),
                    List.of(
                        new WorkbookOperation.EnsureSheet("Budget"),
                        new WorkbookOperation.SetRange(
                            "Budget",
                            "A1:B5",
                            List.of(
                                List.of(new CellInput.Text("Status"), new CellInput.Text("Amount")),
                                List.of(new CellInput.Text("Queued"), new CellInput.Numeric(1.0)),
                                List.of(new CellInput.Text("Done"), new CellInput.Numeric(9.0)),
                                List.of(new CellInput.Text("Done"), new CellInput.Numeric(11.0)),
                                List.of(new CellInput.Text("Queued"), new CellInput.Numeric(4.0)))),
                        new WorkbookOperation.SetConditionalFormatting(
                            "Budget",
                            new ConditionalFormattingBlockInput(
                                List.of("B2:B5"),
                                List.of(
                                    new ConditionalFormattingRuleInput.FormulaRule(
                                        "B2>5",
                                        true,
                                        new DifferentialStyleInput(
                                            "0.00", true, null, null, "#102030", null, null,
                                            "#E0F0AA", null)),
                                    new ConditionalFormattingRuleInput.CellValueRule(
                                        ExcelComparisonOperator.BETWEEN,
                                        "1",
                                        "10",
                                        false,
                                        new DifferentialStyleInput(
                                            null, null, true, null, null, null, null, null,
                                            null)))))),
                    new WorkbookReadOperation.GetConditionalFormatting(
                        "conditional-formatting", "Budget", new RangeSelection.All()),
                    new WorkbookReadOperation.AnalyzeConditionalFormattingHealth(
                        "conditional-formatting-health",
                        new SheetSelection.Selected(List.of("Budget")))));

    GridGrindResponse.Success success = success(response);
    WorkbookReadResult.ConditionalFormattingResult conditionalFormatting =
        read(
            success,
            "conditional-formatting",
            WorkbookReadResult.ConditionalFormattingResult.class);
    WorkbookReadResult.ConditionalFormattingHealthResult health =
        read(
            success,
            "conditional-formatting-health",
            WorkbookReadResult.ConditionalFormattingHealthResult.class);

    assertEquals(workbookPath.toAbsolutePath().toString(), savedPath(success));
    assertEquals(1, conditionalFormatting.conditionalFormattingBlocks().size());
    ConditionalFormattingEntryReport block =
        conditionalFormatting.conditionalFormattingBlocks().getFirst();
    assertEquals(List.of("B2:B5"), block.ranges());
    assertEquals(2, block.rules().size());
    assertInstanceOf(ConditionalFormattingRuleReport.FormulaRule.class, block.rules().get(0));
    assertInstanceOf(ConditionalFormattingRuleReport.CellValueRule.class, block.rules().get(1));
    assertEquals(1, health.analysis().checkedConditionalFormattingBlockCount());
    assertEquals(List.of(), health.analysis().findings());

    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      WorkbookReadExecutor readExecutor = new WorkbookReadExecutor();
      dev.erst.gridgrind.excel.WorkbookReadResult.ConditionalFormattingResult reopened =
          (dev.erst.gridgrind.excel.WorkbookReadResult.ConditionalFormattingResult)
              readExecutor
                  .apply(
                      workbook,
                      new WorkbookReadCommand.GetConditionalFormatting(
                          "conditional-formatting",
                          "Budget",
                          new dev.erst.gridgrind.excel.ExcelRangeSelection.All()))
                  .getFirst();

      assertEquals(1, reopened.conditionalFormattingBlocks().size());
      ExcelConditionalFormattingBlockSnapshot reopenedBlock =
          reopened.conditionalFormattingBlocks().getFirst();
      assertEquals(List.of("B2:B5"), reopenedBlock.ranges());
      assertEquals(2, reopenedBlock.rules().size());
      assertEquals(
          "B2>5",
          cast(
                  ExcelConditionalFormattingRuleSnapshot.FormulaRule.class,
                  reopenedBlock.rules().get(0))
              .formula());
      assertEquals(
          ExcelComparisonOperator.BETWEEN,
          cast(
                  ExcelConditionalFormattingRuleSnapshot.CellValueRule.class,
                  reopenedBlock.rules().get(1))
              .operator());
    }
  }

  @Test
  void returnsAutofilterAndTableReadResultsAndPersistsDefinitions() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-table-autofilter-");

    GridGrindResponse response =
        new DefaultGridGrindRequestExecutor()
            .execute(
                request(
                    new GridGrindRequest.WorkbookSource.New(),
                    new GridGrindRequest.WorkbookPersistence.SaveAs(workbookPath.toString()),
                    List.of(
                        new WorkbookOperation.EnsureSheet("Budget"),
                        new WorkbookOperation.SetRange(
                            "Budget",
                            "A1:C4",
                            List.of(
                                List.of(
                                    new CellInput.Text("Item"),
                                    new CellInput.Text("Amount"),
                                    new CellInput.Text("Billable")),
                                List.of(
                                    new CellInput.Text("Hosting"),
                                    new CellInput.Numeric(49.0),
                                    new CellInput.BooleanValue(true)),
                                List.of(
                                    new CellInput.Text("Domain"),
                                    new CellInput.Numeric(12.0),
                                    new CellInput.BooleanValue(false)),
                                List.of(
                                    new CellInput.Text("Support"),
                                    new CellInput.Numeric(18.0),
                                    new CellInput.BooleanValue(true)))),
                        new WorkbookOperation.SetRange(
                            "Budget",
                            "E1:F3",
                            List.of(
                                List.of(new CellInput.Text("Queue"), new CellInput.Text("Owner")),
                                List.of(
                                    new CellInput.Text("Late invoices"),
                                    new CellInput.Text("Marta")),
                                List.of(
                                    new CellInput.Text("Badge orders"),
                                    new CellInput.Text("Rihards")))),
                        new WorkbookOperation.SetAutofilter("Budget", "E1:F3"),
                        new WorkbookOperation.SetTable(
                            new TableInput(
                                "BudgetTable",
                                "Budget",
                                "A1:C4",
                                false,
                                new TableStyleInput.Named(
                                    "TableStyleMedium2", false, false, true, false)))),
                    new WorkbookReadOperation.GetAutofilters("filters", "Budget"),
                    new WorkbookReadOperation.GetTables("tables", new TableSelection.All()),
                    new WorkbookReadOperation.AnalyzeAutofilterHealth(
                        "autofilter-health", new SheetSelection.Selected(List.of("Budget"))),
                    new WorkbookReadOperation.AnalyzeTableHealth(
                        "table-health", new TableSelection.All())));

    GridGrindResponse.Success success = success(response);
    WorkbookReadResult.AutofiltersResult filters =
        read(success, "filters", WorkbookReadResult.AutofiltersResult.class);
    WorkbookReadResult.TablesResult tables =
        read(success, "tables", WorkbookReadResult.TablesResult.class);
    WorkbookReadResult.AutofilterHealthResult autofilterHealth =
        read(success, "autofilter-health", WorkbookReadResult.AutofilterHealthResult.class);
    WorkbookReadResult.TableHealthResult tableHealth =
        read(success, "table-health", WorkbookReadResult.TableHealthResult.class);

    assertEquals(workbookPath.toAbsolutePath().toString(), savedPath(success));
    assertEquals(
        List.of(
            new AutofilterEntryReport.SheetOwned("E1:F3"),
            new AutofilterEntryReport.TableOwned("A1:C4", "BudgetTable")),
        filters.autofilters());
    assertEquals(1, tables.tables().size());
    TableEntryReport table = tables.tables().getFirst();
    assertEquals("BudgetTable", table.name());
    assertEquals(List.of("Item", "Amount", "Billable"), table.columnNames());
    assertEquals(
        new TableStyleReport.Named("TableStyleMedium2", false, false, true, false), table.style());
    assertTrue(table.hasAutofilter());
    assertEquals(2, autofilterHealth.analysis().checkedAutofilterCount());
    assertEquals(List.of(), autofilterHealth.analysis().findings());
    assertEquals(1, tableHealth.analysis().checkedTableCount());
    assertEquals(List.of(), tableHealth.analysis().findings());

    try (ExcelWorkbook workbook = ExcelWorkbook.open(workbookPath)) {
      WorkbookReadExecutor readExecutor = new WorkbookReadExecutor();
      dev.erst.gridgrind.excel.WorkbookReadResult.AutofiltersResult reopenedFilters =
          (dev.erst.gridgrind.excel.WorkbookReadResult.AutofiltersResult)
              readExecutor
                  .apply(workbook, new WorkbookReadCommand.GetAutofilters("filters", "Budget"))
                  .getFirst();
      dev.erst.gridgrind.excel.WorkbookReadResult.TablesResult reopenedTables =
          (dev.erst.gridgrind.excel.WorkbookReadResult.TablesResult)
              readExecutor
                  .apply(
                      workbook,
                      new WorkbookReadCommand.GetTables("tables", new ExcelTableSelection.All()))
                  .getFirst();

      assertEquals(
          List.of(
              new ExcelAutofilterSnapshot.SheetOwned("E1:F3"),
              new ExcelAutofilterSnapshot.TableOwned("A1:C4", "BudgetTable")),
          reopenedFilters.autofilters());
      assertEquals(
          List.of(
              new ExcelTableSnapshot(
                  "BudgetTable",
                  "Budget",
                  "A1:C4",
                  1,
                  0,
                  List.of("Item", "Amount", "Billable"),
                  new dev.erst.gridgrind.excel.ExcelTableStyleSnapshot.Named(
                      "TableStyleMedium2", false, false, true, false),
                  true)),
          reopenedTables.tables());
    }
  }

  @Test
  void returnsTableReadResultsForByNameSelectionAndNoneStyle() {
    GridGrindResponse response =
        new DefaultGridGrindRequestExecutor()
            .execute(
                request(
                    new GridGrindRequest.WorkbookSource.New(),
                    new GridGrindRequest.WorkbookPersistence.None(),
                    List.of(
                        new WorkbookOperation.EnsureSheet("Budget"),
                        new WorkbookOperation.SetRange(
                            "Budget",
                            "A1:B2",
                            List.of(
                                List.of(new CellInput.Text("Item"), new CellInput.Text("Amount")),
                                List.of(
                                    new CellInput.Text("Hosting"), new CellInput.Numeric(49.0)))),
                        new WorkbookOperation.SetTable(
                            new TableInput(
                                "BudgetTable",
                                "Budget",
                                "A1:B2",
                                false,
                                new TableStyleInput.None()))),
                    new WorkbookReadOperation.GetTables(
                        "tables", new TableSelection.ByNames(List.of("BudgetTable")))));

    WorkbookReadResult.TablesResult tables =
        read(success(response), "tables", WorkbookReadResult.TablesResult.class);

    assertEquals(1, tables.tables().size());
    assertEquals(new TableStyleReport.None(), tables.tables().getFirst().style());
  }

  @Test
  void returnsUnsupportedDataValidationReadEntries() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-data-validation-unsupported-");

    try (XSSFWorkbook workbook = new XSSFWorkbook()) {
      workbook.createSheet("Budget");
      CTDataValidation validation =
          workbook
              .getSheet("Budget")
              .getCTWorksheet()
              .addNewDataValidations()
              .addNewDataValidation();
      validation.setType(STDataValidationType.NONE);
      validation.setSqref(List.of("A1"));
      workbook.getSheet("Budget").getCTWorksheet().getDataValidations().setCount(1L);
      try (var outputStream = Files.newOutputStream(workbookPath)) {
        workbook.write(outputStream);
      }
    }

    GridGrindResponse response =
        new DefaultGridGrindRequestExecutor()
            .execute(
                request(
                    new GridGrindRequest.WorkbookSource.ExistingFile(workbookPath.toString()),
                    new GridGrindRequest.WorkbookPersistence.None(),
                    List.of(),
                    new WorkbookReadOperation.GetDataValidations(
                        "validations", "Budget", new RangeSelection.All())));

    WorkbookReadResult.DataValidationsResult validations =
        read(success(response), "validations", WorkbookReadResult.DataValidationsResult.class);
    DataValidationEntryReport.Unsupported unsupported =
        assertInstanceOf(
            DataValidationEntryReport.Unsupported.class, validations.validations().getFirst());

    assertEquals(List.of("A1"), unsupported.ranges());
    assertEquals("ANY", unsupported.kind());
  }

  @Test
  void returnsFileHyperlinksInCanonicalPathShape() throws IOException {
    Path linkedFileDirectory = Files.createTempDirectory("gridgrind file links ");
    Path linkedFile = linkedFileDirectory.resolve("memo final.pdf");
    Files.writeString(linkedFile, "seed");

    GridGrindResponse response =
        new DefaultGridGrindRequestExecutor()
            .execute(
                request(
                    new GridGrindRequest.WorkbookSource.New(),
                    new GridGrindRequest.WorkbookPersistence.None(),
                    List.of(
                        new WorkbookOperation.EnsureSheet("Budget"),
                        new WorkbookOperation.SetCell("Budget", "A1", new CellInput.Text("Memo")),
                        new WorkbookOperation.SetHyperlink(
                            "Budget", "A1", new HyperlinkTarget.File(linkedFile.toString()))),
                    new WorkbookReadOperation.GetCells("cells", "Budget", List.of("A1")),
                    new WorkbookReadOperation.GetHyperlinks(
                        "hyperlinks", "Budget", new CellSelection.Selected(List.of("A1")))));

    GridGrindResponse.Success success = success(response);
    GridGrindResponse.CellReport.TextReport linkedCell =
        cast(
            GridGrindResponse.CellReport.TextReport.class,
            read(success, "cells", WorkbookReadResult.CellsResult.class).cells().getFirst());
    WorkbookReadResult.HyperlinksResult hyperlinks =
        read(success, "hyperlinks", WorkbookReadResult.HyperlinksResult.class);

    assertEquals(new HyperlinkTarget.File(linkedFile.toString()), linkedCell.hyperlink());
    assertEquals(
        new HyperlinkTarget.File(linkedFile.toString()),
        hyperlinks.hyperlinks().getFirst().hyperlink());
  }

  @Test
  void convertsEmailAndDocumentHyperlinksToCanonicalProtocolTargets() {
    assertEquals(
        new HyperlinkTarget.Email("team@example.com"),
        WorkbookReadResultConverter.toHyperlinkTarget(
            new ExcelHyperlink.Email("team@example.com")));
    assertEquals(
        new HyperlinkTarget.Document("Budget!B4"),
        WorkbookReadResultConverter.toHyperlinkTarget(new ExcelHyperlink.Document("Budget!B4")));
  }

  @Test
  void returnsFactualAndAnalysisReadResults() {
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
                    new WorkbookReadOperation.GetFormulaSurface(
                        "formula", new SheetSelection.All()),
                    new WorkbookReadOperation.GetSheetSchema("schema", "Budget", "A1", 4, 2),
                    new WorkbookReadOperation.GetNamedRangeSurface(
                        "ranges", new NamedRangeSelection.All()),
                    new WorkbookReadOperation.AnalyzeFormulaHealth(
                        "formula-health", new SheetSelection.All())));

    GridGrindResponse.Success success = success(response);
    GridGrindResponse.FormulaSurfaceReport formula =
        read(success, "formula", WorkbookReadResult.FormulaSurfaceResult.class).analysis();
    GridGrindResponse.SheetSchemaReport schema =
        read(success, "schema", WorkbookReadResult.SheetSchemaResult.class).analysis();
    GridGrindResponse.NamedRangeSurfaceReport ranges =
        read(success, "ranges", WorkbookReadResult.NamedRangeSurfaceResult.class).analysis();
    GridGrindResponse.FormulaHealthReport formulaHealth =
        read(success, "formula-health", WorkbookReadResult.FormulaHealthResult.class).analysis();

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
    assertEquals(1, formulaHealth.checkedFormulaCellCount());
    assertEquals(0, formulaHealth.summary().totalCount());
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
    assertEquals(
        "targetIndex out of range: workbook has 1 sheet(s), valid positions are 0 to 0; got 1",
        failure.problem().message());
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
  void returnsStructuredFailureWhenSetSheetPaneTargetsMissingSheet() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new GridGrindRequest.WorkbookSource.New(),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        List.of(
                            new WorkbookOperation.SetSheetPane(
                                "Missing", new PaneInput.Frozen(1, 1, 1, 1))))));

    assertEquals(GridGrindProblemCode.SHEET_NOT_FOUND, failure.problem().code());
    assertEquals("SET_SHEET_PANE", failure.problem().context().operationType());
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
    assertEquals(new ExcelSheetPane.Frozen(1, 2, 3, 4), XlsxRoundTrip.pane(workbookPath, "Budget"));
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
                        List.of(),
                        new WorkbookReadOperation.GetWindow("cells", "Budget", "A1", 5, 5))));

    assertEquals(GridGrindProblemCode.SHEET_NOT_FOUND, failure.problem().code());
    assertEquals("EXECUTE_READ", failure.problem().context().stage());
    assertEquals("GET_WINDOW", failure.problem().context().readType());
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
  void returnsStructuredFailureWhenDeletingLastSheet() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new GridGrindRequest.WorkbookSource.New(),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        List.of(
                            new WorkbookOperation.EnsureSheet("Budget"),
                            new WorkbookOperation.DeleteSheet("Budget")))));

    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertEquals("APPLY_OPERATION", failure.problem().context().stage());
    assertEquals(1, failure.problem().context().operationIndex());
    assertEquals("DELETE_SHEET", failure.problem().context().operationType());
    assertTrue(failure.problem().message().contains("at least one sheet"));
  }

  @Test
  void returnsStructuredFailureWhenDeletingLastVisibleSheet() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new GridGrindRequest.WorkbookSource.New(),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        List.of(
                            new WorkbookOperation.EnsureSheet("Alpha"),
                            new WorkbookOperation.EnsureSheet("Beta"),
                            new WorkbookOperation.SetSheetVisibility(
                                "Beta", ExcelSheetVisibility.HIDDEN),
                            new WorkbookOperation.DeleteSheet("Alpha")))));

    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertEquals("APPLY_OPERATION", failure.problem().context().stage());
    assertEquals(3, failure.problem().context().operationIndex());
    assertEquals("DELETE_SHEET", failure.problem().context().operationType());
    assertEquals("Alpha", failure.problem().context().sheetName());
    assertTrue(failure.problem().message().contains("last visible sheet"));
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
                        List.of(),
                        new WorkbookReadOperation.GetCells("cells", "Missing", List.of("A1")))));

    assertEquals(GridGrindProblemCode.SHEET_NOT_FOUND, failure.problem().code());
    assertFalse(Files.exists(workbookPath));
  }

  @Test
  void returnsInvalidCellAddressForGetCellsWithMalformedAddress() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new GridGrindRequest.WorkbookSource.New(),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        List.of(new WorkbookOperation.EnsureSheet("Data")),
                        new WorkbookReadOperation.GetCells(
                            "cells", "Data", List.of("A1", "BADADDR")))));

    assertEquals(GridGrindProblemCode.INVALID_CELL_ADDRESS, failure.problem().code());
    assertEquals("EXECUTE_READ", failure.problem().context().stage());
    assertEquals("BADADDR", failure.problem().context().address());
  }

  @Test
  void returnsInvalidCellAddressForGetCellsWithZeroRowAddress() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new GridGrindRequest.WorkbookSource.New(),
                        new GridGrindRequest.WorkbookPersistence.None(),
                        List.of(new WorkbookOperation.EnsureSheet("Data")),
                        new WorkbookReadOperation.GetCells("cells", "Data", List.of("A0")))));

    assertEquals(GridGrindProblemCode.INVALID_CELL_ADDRESS, failure.problem().code());
    assertEquals("EXECUTE_READ", failure.problem().context().stage());
    assertEquals("A0", failure.problem().context().address());
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
                            new WorkbookOperation.EnsureSheet("Data"),
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
                            new WorkbookOperation.EnsureSheet("Data"),
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
    assertEquals(
        GridGrindProblemCode.INVALID_FORMULA, failure.problem().causes().getFirst().code());
    assertEquals("APPLY_OPERATION", failure.problem().causes().getFirst().stage());
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
                            new WorkbookOperation.EnsureSheet("Data"),
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
    assertEquals(
        GridGrindProblemCode.SHEET_NOT_FOUND, failure.problem().causes().getFirst().code());
    assertTrue(
        failure
            .problem()
            .causes()
            .get(1)
            .message()
            .contains("Workbook close failed after the primary problem"));
    assertEquals(GridGrindProblemCode.IO_ERROR, failure.problem().causes().get(1).code());
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
                                    new CellAlignmentInput(
                                        true,
                                        ExcelHorizontalAlignment.CENTER,
                                        ExcelVerticalAlignment.CENTER,
                                        null,
                                        null),
                                    new CellFontInput(true, null, null, null, null, null, null),
                                    null,
                                    null,
                                    null)),
                            new WorkbookOperation.ApplyStyle(
                                "Budget",
                                "C1",
                                new CellStyleInput(
                                    null,
                                    new CellAlignmentInput(
                                        null,
                                        ExcelHorizontalAlignment.RIGHT,
                                        ExcelVerticalAlignment.BOTTOM,
                                        null,
                                        null),
                                    new CellFontInput(null, true, null, null, null, null, null),
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
        ExcelHorizontalAlignment.CENTER,
        cells.cells().getFirst().style().alignment().horizontalAlignment());
    assertTrue(cells.cells().getFirst().style().font().bold());
    assertTrue(cells.cells().getFirst().style().alignment().wrapText());
    assertEquals("BLANK", cells.cells().get(1).effectiveType());
    assertEquals("49", cells.cells().get(2).displayValue());
    assertTrue(
        window.rows().getFirst().cells().stream().anyMatch(cell -> "C1".equals(cell.address())));
    assertTrue(cells.cells().get(3).style().font().italic());
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
                                    new CellAlignmentInput(
                                        true,
                                        ExcelHorizontalAlignment.CENTER,
                                        ExcelVerticalAlignment.TOP,
                                        45,
                                        3),
                                    new CellFontInput(
                                        true,
                                        false,
                                        "Aptos",
                                        new FontHeightInput.Points(new BigDecimal("11.5")),
                                        "#1F4E78",
                                        true,
                                        true),
                                    new CellFillInput(
                                        dev.erst.gridgrind.excel.ExcelFillPattern
                                            .THIN_HORIZONTAL_BANDS,
                                        "#FFF2CC",
                                        "#DDEBF7"),
                                    new CellBorderInput(
                                        new CellBorderSideInput(ExcelBorderStyle.THIN, "#102030"),
                                        null,
                                        new CellBorderSideInput(ExcelBorderStyle.DOUBLE, "#203040"),
                                        null,
                                        null),
                                    new CellProtectionInput(false, true)))),
                        new WorkbookReadOperation.GetCells("cells", "Budget", List.of("A1")))));

    GridGrindResponse.CellStyleReport style =
        read(success, "cells", WorkbookReadResult.CellsResult.class).cells().getFirst().style();

    assertTrue(Files.exists(workbookPath));
    assertTrue(style.font().bold());
    assertFalse(style.font().italic());
    assertTrue(style.alignment().wrapText());
    assertEquals(ExcelHorizontalAlignment.CENTER, style.alignment().horizontalAlignment());
    assertEquals(ExcelVerticalAlignment.TOP, style.alignment().verticalAlignment());
    assertEquals(45, style.alignment().textRotation());
    assertEquals(3, style.alignment().indentation());
    assertEquals("Aptos", style.font().fontName());
    assertEquals(230, style.font().fontHeight().twips());
    assertEquals(new BigDecimal("11.5"), style.font().fontHeight().points());
    assertEquals("#1F4E78", style.font().fontColor());
    assertTrue(style.font().underline());
    assertTrue(style.font().strikeout());
    assertEquals(
        dev.erst.gridgrind.excel.ExcelFillPattern.THIN_HORIZONTAL_BANDS, style.fill().pattern());
    assertEquals("#FFF2CC", style.fill().foregroundColor());
    assertEquals("#DDEBF7", style.fill().backgroundColor());
    assertEquals(ExcelBorderStyle.THIN, style.border().top().style());
    assertEquals(ExcelBorderStyle.DOUBLE, style.border().right().style());
    assertEquals(ExcelBorderStyle.THIN, style.border().bottom().style());
    assertEquals(ExcelBorderStyle.THIN, style.border().left().style());
    assertEquals("#102030", style.border().top().color());
    assertEquals("#203040", style.border().right().color());
    assertFalse(style.protection().locked());
    assertTrue(style.protection().hiddenFormula());
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
        Path.of("/tmp/out.xlsx").toAbsolutePath().normalize().toString(),
        executor.persistencePath(newSource, saveAs));
    assertEquals(
        Path.of("/tmp/source.xlsx").toAbsolutePath().normalize().toString(),
        executor.persistencePath(existingFile, overwrite));
    assertNull(executor.persistencePath(newSource, overwrite));
    assertNull(executor.persistencePath(newSource, none));
    assertNull(executor.persistencePath(existingFile, none));
  }

  @Test
  void persistWorkbookRejectsOverwriteForNewSources() throws Exception {
    DefaultGridGrindRequestExecutor executor = new DefaultGridGrindRequestExecutor();

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      IllegalArgumentException exception =
          assertThrows(
              IllegalArgumentException.class,
              () ->
                  executor.persistWorkbook(
                      workbook,
                      new GridGrindRequest.WorkbookSource.New(),
                      new GridGrindRequest.WorkbookPersistence.OverwriteSource()));

      assertEquals("OVERWRITE persistence requires an EXISTING source", exception.getMessage());
    }
  }

  @Test
  void persistencePathNormalizesDoubleDotSegments() {
    DefaultGridGrindRequestExecutor executor = new DefaultGridGrindRequestExecutor();
    GridGrindRequest.WorkbookSource newSource = new GridGrindRequest.WorkbookSource.New();
    GridGrindRequest.WorkbookPersistence saveAs =
        new GridGrindRequest.WorkbookPersistence.SaveAs("/tmp/subdir/../out.xlsx");

    assertEquals("/tmp/out.xlsx", executor.persistencePath(newSource, saveAs));
  }

  @Test
  void persistWorkbookSaveAsReportsNormalizedExecutionPath() throws Exception {
    DefaultGridGrindRequestExecutor executor = new DefaultGridGrindRequestExecutor();
    Path tempDir = Files.createTempDirectory("gridgrind-normalize-test-");
    Path subDir = Files.createDirectory(tempDir.resolve("subdir"));
    String pathWithDotDot = subDir + "/../out.xlsx";

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      GridGrindResponse.PersistenceOutcome outcome =
          executor.persistWorkbook(
              workbook,
              new GridGrindRequest.WorkbookSource.New(),
              new GridGrindRequest.WorkbookPersistence.SaveAs(pathWithDotDot));

      GridGrindResponse.PersistenceOutcome.SavedAs savedAs =
          assertInstanceOf(GridGrindResponse.PersistenceOutcome.SavedAs.class, outcome);
      assertEquals(pathWithDotDot, savedAs.requestedPath());
      assertEquals(tempDir.resolve("out.xlsx").toString(), savedAs.executionPath());
    } finally {
      Files.deleteIfExists(tempDir.resolve("out.xlsx"));
      Files.deleteIfExists(subDir);
      Files.deleteIfExists(tempDir);
    }
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
                                    null,
                                    null,
                                    new CellFontInput(true, null, null, null, null, null, null),
                                    null,
                                    null,
                                    null))))));

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
            new WorkbookOperation.SetCell("S", "A1", new CellInput.Text("hello")), exception));
    assertNull(
        DefaultGridGrindRequestExecutor.formulaFor(
            new WorkbookOperation.SetCell(
                "S",
                "A1",
                new CellInput.RichText(
                    List.of(
                        new RichTextRunInput("Budget", null),
                        new RichTextRunInput(
                            " FY26",
                            new CellFontInput(true, false, null, null, "#AABBCC", false, false))))),
            exception));
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
  void extractsSheetStateContextForB1Operations() {
    RuntimeException exception = new RuntimeException("test");
    WorkbookOperation copySheet =
        new WorkbookOperation.CopySheet(
            "Budget", "Budget Copy", new SheetCopyPosition.AppendAtEnd());
    WorkbookOperation setActiveSheet = new WorkbookOperation.SetActiveSheet("Budget Copy");
    WorkbookOperation setSelectedSheets =
        new WorkbookOperation.SetSelectedSheets(List.of("Budget", "Budget Copy"));
    WorkbookOperation setSheetVisibility =
        new WorkbookOperation.SetSheetVisibility("Budget", ExcelSheetVisibility.HIDDEN);
    WorkbookOperation setSheetProtection =
        new WorkbookOperation.SetSheetProtection("Budget", protectionSettings());
    WorkbookOperation clearSheetProtection = new WorkbookOperation.ClearSheetProtection("Budget");

    assertNull(DefaultGridGrindRequestExecutor.formulaFor(copySheet, exception));
    assertNull(DefaultGridGrindRequestExecutor.formulaFor(setActiveSheet, exception));
    assertNull(DefaultGridGrindRequestExecutor.formulaFor(setSelectedSheets, exception));
    assertNull(DefaultGridGrindRequestExecutor.formulaFor(setSheetVisibility, exception));
    assertNull(DefaultGridGrindRequestExecutor.formulaFor(setSheetProtection, exception));
    assertNull(DefaultGridGrindRequestExecutor.formulaFor(clearSheetProtection, exception));

    assertEquals("Budget", DefaultGridGrindRequestExecutor.sheetNameFor(copySheet, exception));
    assertEquals(
        "Budget Copy", DefaultGridGrindRequestExecutor.sheetNameFor(setActiveSheet, exception));
    assertNull(DefaultGridGrindRequestExecutor.sheetNameFor(setSelectedSheets, exception));
    assertEquals(
        "Budget", DefaultGridGrindRequestExecutor.sheetNameFor(setSheetVisibility, exception));
    assertEquals(
        "Budget", DefaultGridGrindRequestExecutor.sheetNameFor(setSheetProtection, exception));
    assertEquals(
        "Budget", DefaultGridGrindRequestExecutor.sheetNameFor(clearSheetProtection, exception));

    assertNull(DefaultGridGrindRequestExecutor.addressFor(copySheet, exception));
    assertNull(DefaultGridGrindRequestExecutor.addressFor(setActiveSheet, exception));
    assertNull(DefaultGridGrindRequestExecutor.addressFor(setSelectedSheets, exception));
    assertNull(DefaultGridGrindRequestExecutor.addressFor(setSheetVisibility, exception));
    assertNull(DefaultGridGrindRequestExecutor.addressFor(setSheetProtection, exception));
    assertNull(DefaultGridGrindRequestExecutor.addressFor(clearSheetProtection, exception));

    assertNull(DefaultGridGrindRequestExecutor.rangeFor(copySheet, exception));
    assertNull(DefaultGridGrindRequestExecutor.rangeFor(setActiveSheet, exception));
    assertNull(DefaultGridGrindRequestExecutor.rangeFor(setSelectedSheets, exception));
    assertNull(DefaultGridGrindRequestExecutor.rangeFor(setSheetVisibility, exception));
    assertNull(DefaultGridGrindRequestExecutor.rangeFor(setSheetProtection, exception));
    assertNull(DefaultGridGrindRequestExecutor.rangeFor(clearSheetProtection, exception));

    assertNull(DefaultGridGrindRequestExecutor.namedRangeNameFor(copySheet, exception));
    assertNull(DefaultGridGrindRequestExecutor.namedRangeNameFor(setActiveSheet, exception));
    assertNull(DefaultGridGrindRequestExecutor.namedRangeNameFor(setSelectedSheets, exception));
    assertNull(DefaultGridGrindRequestExecutor.namedRangeNameFor(setSheetVisibility, exception));
    assertNull(DefaultGridGrindRequestExecutor.namedRangeNameFor(setSheetProtection, exception));
    assertNull(DefaultGridGrindRequestExecutor.namedRangeNameFor(clearSheetProtection, exception));
  }

  @Test
  @SuppressWarnings("PMD.NcssCount")
  void extractsContextForStructuralLayoutOperations() {
    RuntimeException exception = new RuntimeException("test");
    WorkbookOperation mergeCells = new WorkbookOperation.MergeCells("Budget", "A1:B2");
    WorkbookOperation unmergeCells = new WorkbookOperation.UnmergeCells("Budget", "A1:B2");
    WorkbookOperation setColumnWidth = new WorkbookOperation.SetColumnWidth("Budget", 0, 1, 16.0);
    WorkbookOperation setRowHeight = new WorkbookOperation.SetRowHeight("Budget", 0, 1, 28.5);
    WorkbookOperation insertRows = new WorkbookOperation.InsertRows("Budget", 1, 2);
    WorkbookOperation deleteRows =
        new WorkbookOperation.DeleteRows("Budget", new RowSpanInput.Band(1, 2));
    WorkbookOperation shiftRows =
        new WorkbookOperation.ShiftRows("Budget", new RowSpanInput.Band(1, 2), 1);
    WorkbookOperation insertColumns = new WorkbookOperation.InsertColumns("Budget", 1, 2);
    WorkbookOperation deleteColumns =
        new WorkbookOperation.DeleteColumns("Budget", new ColumnSpanInput.Band(1, 2));
    WorkbookOperation shiftColumns =
        new WorkbookOperation.ShiftColumns("Budget", new ColumnSpanInput.Band(1, 2), -1);
    WorkbookOperation setRowVisibility =
        new WorkbookOperation.SetRowVisibility("Budget", new RowSpanInput.Band(1, 2), true);
    WorkbookOperation setColumnVisibility =
        new WorkbookOperation.SetColumnVisibility("Budget", new ColumnSpanInput.Band(1, 2), false);
    WorkbookOperation groupRows =
        new WorkbookOperation.GroupRows("Budget", new RowSpanInput.Band(1, 2), true);
    WorkbookOperation ungroupRows =
        new WorkbookOperation.UngroupRows("Budget", new RowSpanInput.Band(1, 2));
    WorkbookOperation groupColumns =
        new WorkbookOperation.GroupColumns("Budget", new ColumnSpanInput.Band(1, 2), true);
    WorkbookOperation ungroupColumns =
        new WorkbookOperation.UngroupColumns("Budget", new ColumnSpanInput.Band(1, 2));
    WorkbookOperation setSheetPane =
        new WorkbookOperation.SetSheetPane("Budget", new PaneInput.Frozen(1, 1, 1, 1));
    WorkbookOperation setSheetZoom = new WorkbookOperation.SetSheetZoom("Budget", 125);
    WorkbookOperation setPrintLayout =
        new WorkbookOperation.SetPrintLayout(
            "Budget",
            new PrintLayoutInput(
                new PrintAreaInput.Range("A1:B12"),
                ExcelPrintOrientation.LANDSCAPE,
                new PrintScalingInput.Fit(1, 0),
                new PrintTitleRowsInput.Band(0, 0),
                new PrintTitleColumnsInput.Band(0, 0),
                new HeaderFooterTextInput("Budget", "", ""),
                new HeaderFooterTextInput("", "Page &P", "")));
    WorkbookOperation clearPrintLayout = new WorkbookOperation.ClearPrintLayout("Budget");

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

    assertNull(DefaultGridGrindRequestExecutor.formulaFor(insertRows, exception));
    assertEquals("Budget", DefaultGridGrindRequestExecutor.sheetNameFor(insertRows, exception));
    assertNull(DefaultGridGrindRequestExecutor.addressFor(insertRows, exception));
    assertNull(DefaultGridGrindRequestExecutor.rangeFor(insertRows, exception));

    assertNull(DefaultGridGrindRequestExecutor.formulaFor(deleteRows, exception));
    assertEquals("Budget", DefaultGridGrindRequestExecutor.sheetNameFor(deleteRows, exception));
    assertNull(DefaultGridGrindRequestExecutor.addressFor(deleteRows, exception));
    assertNull(DefaultGridGrindRequestExecutor.rangeFor(deleteRows, exception));

    assertNull(DefaultGridGrindRequestExecutor.formulaFor(shiftRows, exception));
    assertEquals("Budget", DefaultGridGrindRequestExecutor.sheetNameFor(shiftRows, exception));
    assertNull(DefaultGridGrindRequestExecutor.addressFor(shiftRows, exception));
    assertNull(DefaultGridGrindRequestExecutor.rangeFor(shiftRows, exception));

    assertNull(DefaultGridGrindRequestExecutor.formulaFor(insertColumns, exception));
    assertEquals("Budget", DefaultGridGrindRequestExecutor.sheetNameFor(insertColumns, exception));
    assertNull(DefaultGridGrindRequestExecutor.addressFor(insertColumns, exception));
    assertNull(DefaultGridGrindRequestExecutor.rangeFor(insertColumns, exception));

    assertNull(DefaultGridGrindRequestExecutor.formulaFor(deleteColumns, exception));
    assertEquals("Budget", DefaultGridGrindRequestExecutor.sheetNameFor(deleteColumns, exception));
    assertNull(DefaultGridGrindRequestExecutor.addressFor(deleteColumns, exception));
    assertNull(DefaultGridGrindRequestExecutor.rangeFor(deleteColumns, exception));

    assertNull(DefaultGridGrindRequestExecutor.formulaFor(shiftColumns, exception));
    assertEquals("Budget", DefaultGridGrindRequestExecutor.sheetNameFor(shiftColumns, exception));
    assertNull(DefaultGridGrindRequestExecutor.addressFor(shiftColumns, exception));
    assertNull(DefaultGridGrindRequestExecutor.rangeFor(shiftColumns, exception));

    assertNull(DefaultGridGrindRequestExecutor.formulaFor(setRowVisibility, exception));
    assertEquals(
        "Budget", DefaultGridGrindRequestExecutor.sheetNameFor(setRowVisibility, exception));
    assertNull(DefaultGridGrindRequestExecutor.addressFor(setRowVisibility, exception));
    assertNull(DefaultGridGrindRequestExecutor.rangeFor(setRowVisibility, exception));

    assertNull(DefaultGridGrindRequestExecutor.formulaFor(setColumnVisibility, exception));
    assertEquals(
        "Budget", DefaultGridGrindRequestExecutor.sheetNameFor(setColumnVisibility, exception));
    assertNull(DefaultGridGrindRequestExecutor.addressFor(setColumnVisibility, exception));
    assertNull(DefaultGridGrindRequestExecutor.rangeFor(setColumnVisibility, exception));

    assertNull(DefaultGridGrindRequestExecutor.formulaFor(groupRows, exception));
    assertEquals("Budget", DefaultGridGrindRequestExecutor.sheetNameFor(groupRows, exception));
    assertNull(DefaultGridGrindRequestExecutor.addressFor(groupRows, exception));
    assertNull(DefaultGridGrindRequestExecutor.rangeFor(groupRows, exception));

    assertNull(DefaultGridGrindRequestExecutor.formulaFor(ungroupRows, exception));
    assertEquals("Budget", DefaultGridGrindRequestExecutor.sheetNameFor(ungroupRows, exception));
    assertNull(DefaultGridGrindRequestExecutor.addressFor(ungroupRows, exception));
    assertNull(DefaultGridGrindRequestExecutor.rangeFor(ungroupRows, exception));

    assertNull(DefaultGridGrindRequestExecutor.formulaFor(groupColumns, exception));
    assertEquals("Budget", DefaultGridGrindRequestExecutor.sheetNameFor(groupColumns, exception));
    assertNull(DefaultGridGrindRequestExecutor.addressFor(groupColumns, exception));
    assertNull(DefaultGridGrindRequestExecutor.rangeFor(groupColumns, exception));

    assertNull(DefaultGridGrindRequestExecutor.formulaFor(ungroupColumns, exception));
    assertEquals("Budget", DefaultGridGrindRequestExecutor.sheetNameFor(ungroupColumns, exception));
    assertNull(DefaultGridGrindRequestExecutor.addressFor(ungroupColumns, exception));
    assertNull(DefaultGridGrindRequestExecutor.rangeFor(ungroupColumns, exception));

    assertNull(DefaultGridGrindRequestExecutor.formulaFor(setSheetPane, exception));
    assertEquals("Budget", DefaultGridGrindRequestExecutor.sheetNameFor(setSheetPane, exception));
    assertNull(DefaultGridGrindRequestExecutor.addressFor(setSheetPane, exception));
    assertNull(DefaultGridGrindRequestExecutor.rangeFor(setSheetPane, exception));

    assertNull(DefaultGridGrindRequestExecutor.formulaFor(setSheetZoom, exception));
    assertEquals("Budget", DefaultGridGrindRequestExecutor.sheetNameFor(setSheetZoom, exception));
    assertNull(DefaultGridGrindRequestExecutor.addressFor(setSheetZoom, exception));
    assertNull(DefaultGridGrindRequestExecutor.rangeFor(setSheetZoom, exception));

    assertNull(DefaultGridGrindRequestExecutor.formulaFor(setPrintLayout, exception));
    assertEquals("Budget", DefaultGridGrindRequestExecutor.sheetNameFor(setPrintLayout, exception));
    assertNull(DefaultGridGrindRequestExecutor.addressFor(setPrintLayout, exception));
    assertNull(DefaultGridGrindRequestExecutor.rangeFor(setPrintLayout, exception));

    assertNull(DefaultGridGrindRequestExecutor.formulaFor(clearPrintLayout, exception));
    assertEquals(
        "Budget", DefaultGridGrindRequestExecutor.sheetNameFor(clearPrintLayout, exception));
    assertNull(DefaultGridGrindRequestExecutor.addressFor(clearPrintLayout, exception));
    assertNull(DefaultGridGrindRequestExecutor.rangeFor(clearPrintLayout, exception));
  }

  @Test
  void convertsWaveThreeOperationsIntoWorkbookCommands() {
    WorkbookCommand setHyperlink =
        WorkbookCommandConverter.toCommand(
            new WorkbookOperation.SetHyperlink(
                "Budget", "A1", new HyperlinkTarget.Url("https://example.com/report")));
    WorkbookCommand clearHyperlink =
        WorkbookCommandConverter.toCommand(new WorkbookOperation.ClearHyperlink("Budget", "A1"));
    WorkbookCommand setComment =
        WorkbookCommandConverter.toCommand(
            new WorkbookOperation.SetComment(
                "Budget", "A1", new CommentInput("Review", "GridGrind", false)));
    WorkbookCommand clearComment =
        WorkbookCommandConverter.toCommand(new WorkbookOperation.ClearComment("Budget", "A1"));
    WorkbookCommand setNamedRange =
        WorkbookCommandConverter.toCommand(
            new WorkbookOperation.SetNamedRange(
                "BudgetTotal",
                new NamedRangeScope.Workbook(),
                new NamedRangeTarget("Budget", "B4")));
    WorkbookCommand deleteNamedRange =
        WorkbookCommandConverter.toCommand(
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
        WorkbookCommandConverter.toCommand(new WorkbookOperation.EnsureSheet("Budget")));
    assertInstanceOf(
        WorkbookCommand.RenameSheet.class,
        WorkbookCommandConverter.toCommand(new WorkbookOperation.RenameSheet("Budget", "Summary")));
    assertInstanceOf(
        WorkbookCommand.DeleteSheet.class,
        WorkbookCommandConverter.toCommand(new WorkbookOperation.DeleteSheet("Budget")));
    assertInstanceOf(
        WorkbookCommand.MoveSheet.class,
        WorkbookCommandConverter.toCommand(new WorkbookOperation.MoveSheet("Budget", 0)));
    assertInstanceOf(
        WorkbookCommand.MergeCells.class,
        WorkbookCommandConverter.toCommand(new WorkbookOperation.MergeCells("Budget", "A1:B2")));
    assertInstanceOf(
        WorkbookCommand.UnmergeCells.class,
        WorkbookCommandConverter.toCommand(new WorkbookOperation.UnmergeCells("Budget", "A1:B2")));
    assertInstanceOf(
        WorkbookCommand.SetColumnWidth.class,
        WorkbookCommandConverter.toCommand(
            new WorkbookOperation.SetColumnWidth("Budget", 0, 1, 16.0)));
    assertInstanceOf(
        WorkbookCommand.SetRowHeight.class,
        WorkbookCommandConverter.toCommand(
            new WorkbookOperation.SetRowHeight("Budget", 0, 1, 28.5)));
    assertInstanceOf(
        WorkbookCommand.InsertRows.class,
        WorkbookCommandConverter.toCommand(new WorkbookOperation.InsertRows("Budget", 1, 2)));
    assertInstanceOf(
        WorkbookCommand.DeleteRows.class,
        WorkbookCommandConverter.toCommand(
            new WorkbookOperation.DeleteRows("Budget", new RowSpanInput.Band(1, 2))));
    assertInstanceOf(
        WorkbookCommand.ShiftRows.class,
        WorkbookCommandConverter.toCommand(
            new WorkbookOperation.ShiftRows("Budget", new RowSpanInput.Band(1, 2), 1)));
    assertInstanceOf(
        WorkbookCommand.InsertColumns.class,
        WorkbookCommandConverter.toCommand(new WorkbookOperation.InsertColumns("Budget", 1, 2)));
    assertInstanceOf(
        WorkbookCommand.DeleteColumns.class,
        WorkbookCommandConverter.toCommand(
            new WorkbookOperation.DeleteColumns("Budget", new ColumnSpanInput.Band(1, 2))));
    assertInstanceOf(
        WorkbookCommand.ShiftColumns.class,
        WorkbookCommandConverter.toCommand(
            new WorkbookOperation.ShiftColumns("Budget", new ColumnSpanInput.Band(1, 2), -1)));
    assertInstanceOf(
        WorkbookCommand.SetRowVisibility.class,
        WorkbookCommandConverter.toCommand(
            new WorkbookOperation.SetRowVisibility("Budget", new RowSpanInput.Band(1, 2), true)));
    assertInstanceOf(
        WorkbookCommand.SetColumnVisibility.class,
        WorkbookCommandConverter.toCommand(
            new WorkbookOperation.SetColumnVisibility(
                "Budget", new ColumnSpanInput.Band(1, 2), true)));
    assertInstanceOf(
        WorkbookCommand.GroupRows.class,
        WorkbookCommandConverter.toCommand(
            new WorkbookOperation.GroupRows("Budget", new RowSpanInput.Band(1, 2), false)));
    assertInstanceOf(
        WorkbookCommand.UngroupRows.class,
        WorkbookCommandConverter.toCommand(
            new WorkbookOperation.UngroupRows("Budget", new RowSpanInput.Band(1, 2))));
    assertInstanceOf(
        WorkbookCommand.GroupColumns.class,
        WorkbookCommandConverter.toCommand(
            new WorkbookOperation.GroupColumns("Budget", new ColumnSpanInput.Band(1, 2), false)));
    assertInstanceOf(
        WorkbookCommand.UngroupColumns.class,
        WorkbookCommandConverter.toCommand(
            new WorkbookOperation.UngroupColumns("Budget", new ColumnSpanInput.Band(1, 2))));
    assertInstanceOf(
        WorkbookCommand.SetSheetPane.class,
        WorkbookCommandConverter.toCommand(
            new WorkbookOperation.SetSheetPane("Budget", new PaneInput.Frozen(1, 1, 1, 1))));
    assertInstanceOf(
        WorkbookCommand.SetSheetZoom.class,
        WorkbookCommandConverter.toCommand(new WorkbookOperation.SetSheetZoom("Budget", 125)));
    assertInstanceOf(
        WorkbookCommand.SetPrintLayout.class,
        WorkbookCommandConverter.toCommand(
            new WorkbookOperation.SetPrintLayout(
                "Budget",
                new PrintLayoutInput(
                    new PrintAreaInput.Range("A1:B12"),
                    ExcelPrintOrientation.LANDSCAPE,
                    new PrintScalingInput.Fit(1, 0),
                    new PrintTitleRowsInput.Band(0, 0),
                    new PrintTitleColumnsInput.Band(0, 0),
                    new HeaderFooterTextInput("Budget", "", ""),
                    new HeaderFooterTextInput("", "Page &P", "")))));
    assertInstanceOf(
        WorkbookCommand.ClearPrintLayout.class,
        WorkbookCommandConverter.toCommand(new WorkbookOperation.ClearPrintLayout("Budget")));
    assertInstanceOf(
        WorkbookCommand.SetCell.class,
        WorkbookCommandConverter.toCommand(
            new WorkbookOperation.SetCell("Budget", "A1", new CellInput.Text("x"))));
    assertInstanceOf(
        WorkbookCommand.SetRange.class,
        WorkbookCommandConverter.toCommand(
            new WorkbookOperation.SetRange(
                "Budget", "A1:B1", List.of(List.of(new CellInput.Text("x"))))));
    assertInstanceOf(
        WorkbookCommand.ClearRange.class,
        WorkbookCommandConverter.toCommand(new WorkbookOperation.ClearRange("Budget", "A1:B1")));
    assertInstanceOf(
        WorkbookCommand.SetConditionalFormatting.class,
        WorkbookCommandConverter.toCommand(
            new WorkbookOperation.SetConditionalFormatting(
                "Budget",
                new ConditionalFormattingBlockInput(
                    List.of("A1:A3"),
                    List.of(
                        new ConditionalFormattingRuleInput.FormulaRule(
                            "A1>0",
                            false,
                            new DifferentialStyleInput(
                                null, true, null, null, null, null, null, null, null)))))));
    assertInstanceOf(
        WorkbookCommand.ClearConditionalFormatting.class,
        WorkbookCommandConverter.toCommand(
            new WorkbookOperation.ClearConditionalFormatting("Budget", new RangeSelection.All())));
    assertInstanceOf(
        WorkbookCommand.SetAutofilter.class,
        WorkbookCommandConverter.toCommand(new WorkbookOperation.SetAutofilter("Budget", "A1:B4")));
    assertInstanceOf(
        WorkbookCommand.ClearAutofilter.class,
        WorkbookCommandConverter.toCommand(new WorkbookOperation.ClearAutofilter("Budget")));
    assertInstanceOf(
        WorkbookCommand.SetTable.class,
        WorkbookCommandConverter.toCommand(
            new WorkbookOperation.SetTable(
                new TableInput(
                    "BudgetTable", "Budget", "A1:B4", false, new TableStyleInput.None()))));
    assertInstanceOf(
        WorkbookCommand.DeleteTable.class,
        WorkbookCommandConverter.toCommand(
            new WorkbookOperation.DeleteTable("BudgetTable", "Budget")));
    assertInstanceOf(
        WorkbookCommand.AppendRow.class,
        WorkbookCommandConverter.toCommand(
            new WorkbookOperation.AppendRow("Budget", List.of(new CellInput.Text("x")))));
    assertInstanceOf(
        WorkbookCommand.AutoSizeColumns.class,
        WorkbookCommandConverter.toCommand(new WorkbookOperation.AutoSizeColumns("Budget")));
    assertInstanceOf(
        WorkbookCommand.EvaluateAllFormulas.class,
        WorkbookCommandConverter.toCommand(new WorkbookOperation.EvaluateFormulas()));
    assertInstanceOf(
        WorkbookCommand.ForceFormulaRecalculationOnOpen.class,
        WorkbookCommandConverter.toCommand(
            new WorkbookOperation.ForceFormulaRecalculationOnOpen()));

    WorkbookCommand.SetSheetPane setSheetPaneNone =
        cast(
            WorkbookCommand.SetSheetPane.class,
            WorkbookCommandConverter.toCommand(
                new WorkbookOperation.SetSheetPane("Budget", new PaneInput.None())));
    WorkbookCommand.SetSheetPane setSheetPaneSplit =
        cast(
            WorkbookCommand.SetSheetPane.class,
            WorkbookCommandConverter.toCommand(
                new WorkbookOperation.SetSheetPane(
                    "Budget", new PaneInput.Split(1200, 2400, 3, 4, ExcelPaneRegion.LOWER_RIGHT))));
    WorkbookCommand.SetPrintLayout defaultPrintLayout =
        cast(
            WorkbookCommand.SetPrintLayout.class,
            WorkbookCommandConverter.toCommand(
                new WorkbookOperation.SetPrintLayout(
                    "Budget", new PrintLayoutInput(null, null, null, null, null, null, null))));

    assertEquals(new ExcelSheetPane.None(), setSheetPaneNone.pane());
    assertEquals(
        new ExcelSheetPane.Split(1200, 2400, 3, 4, ExcelPaneRegion.LOWER_RIGHT),
        setSheetPaneSplit.pane());
    assertEquals(
        new dev.erst.gridgrind.excel.ExcelPrintLayout(
            new dev.erst.gridgrind.excel.ExcelPrintLayout.Area.None(),
            ExcelPrintOrientation.PORTRAIT,
            new dev.erst.gridgrind.excel.ExcelPrintLayout.Scaling.Automatic(),
            new dev.erst.gridgrind.excel.ExcelPrintLayout.TitleRows.None(),
            new dev.erst.gridgrind.excel.ExcelPrintLayout.TitleColumns.None(),
            new dev.erst.gridgrind.excel.ExcelHeaderFooterText("", "", ""),
            new dev.erst.gridgrind.excel.ExcelHeaderFooterText("", "", "")),
        defaultPrintLayout.printLayout());
  }

  @Test
  void convertsStructuralAndSurfaceReadOperationsIntoWorkbookReadCommands() {
    WorkbookReadCommand workbookSummary =
        WorkbookReadCommandConverter.toReadCommand(
            new WorkbookReadOperation.GetWorkbookSummary("workbook"));
    WorkbookReadCommand namedRanges =
        WorkbookReadCommandConverter.toReadCommand(
            new WorkbookReadOperation.GetNamedRanges(
                "ranges",
                new NamedRangeSelection.Selected(
                    List.of(
                        new NamedRangeSelector.ByName("BudgetTotal"),
                        new NamedRangeSelector.SheetScope("LocalItem", "Budget")))));
    WorkbookReadCommand sheetSummary =
        WorkbookReadCommandConverter.toReadCommand(
            new WorkbookReadOperation.GetSheetSummary("sheet", "Budget"));
    WorkbookReadCommand cells =
        WorkbookReadCommandConverter.toReadCommand(
            new WorkbookReadOperation.GetCells("cells", "Budget", List.of("A1")));
    WorkbookReadCommand window =
        WorkbookReadCommandConverter.toReadCommand(
            new WorkbookReadOperation.GetWindow("window", "Budget", "A1", 2, 2));
    WorkbookReadCommand merged =
        WorkbookReadCommandConverter.toReadCommand(
            new WorkbookReadOperation.GetMergedRegions("merged", "Budget"));
    WorkbookReadCommand hyperlinks =
        WorkbookReadCommandConverter.toReadCommand(
            new WorkbookReadOperation.GetHyperlinks(
                "hyperlinks", "Budget", new CellSelection.AllUsedCells()));
    WorkbookReadCommand comments =
        WorkbookReadCommandConverter.toReadCommand(
            new WorkbookReadOperation.GetComments(
                "comments", "Budget", new CellSelection.Selected(List.of("A1"))));
    WorkbookReadCommand layout =
        WorkbookReadCommandConverter.toReadCommand(
            new WorkbookReadOperation.GetSheetLayout("layout", "Budget"));
    WorkbookReadCommand printLayout =
        WorkbookReadCommandConverter.toReadCommand(
            new WorkbookReadOperation.GetPrintLayout("printLayout", "Budget"));
    WorkbookReadCommand validations =
        WorkbookReadCommandConverter.toReadCommand(
            new WorkbookReadOperation.GetDataValidations(
                "validations", "Budget", new RangeSelection.All()));
    WorkbookReadCommand conditionalFormatting =
        WorkbookReadCommandConverter.toReadCommand(
            new WorkbookReadOperation.GetConditionalFormatting(
                "conditionalFormatting", "Budget", new RangeSelection.Selected(List.of("A1:A3"))));
    WorkbookReadCommand autofilters =
        WorkbookReadCommandConverter.toReadCommand(
            new WorkbookReadOperation.GetAutofilters("autofilters", "Budget"));
    WorkbookReadCommand tables =
        WorkbookReadCommandConverter.toReadCommand(
            new WorkbookReadOperation.GetTables(
                "tables", new TableSelection.ByNames(List.of("BudgetTable"))));
    WorkbookReadCommand formulaSurface =
        WorkbookReadCommandConverter.toReadCommand(
            new WorkbookReadOperation.GetFormulaSurface(
                "formula", new SheetSelection.Selected(List.of("Budget"))));
    WorkbookReadCommand schema =
        WorkbookReadCommandConverter.toReadCommand(
            new WorkbookReadOperation.GetSheetSchema("schema", "Budget", "A1", 3, 2));
    WorkbookReadCommand namedRangeSurface =
        WorkbookReadCommandConverter.toReadCommand(
            new WorkbookReadOperation.GetNamedRangeSurface(
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
    assertInstanceOf(WorkbookReadCommand.GetPrintLayout.class, printLayout);
    assertInstanceOf(WorkbookReadCommand.GetDataValidations.class, validations);
    assertInstanceOf(WorkbookReadCommand.GetConditionalFormatting.class, conditionalFormatting);
    assertInstanceOf(WorkbookReadCommand.GetAutofilters.class, autofilters);
    assertInstanceOf(WorkbookReadCommand.GetTables.class, tables);
    assertInstanceOf(WorkbookReadCommand.GetFormulaSurface.class, formulaSurface);
    assertInstanceOf(WorkbookReadCommand.GetSheetSchema.class, schema);
    assertInstanceOf(WorkbookReadCommand.GetNamedRangeSurface.class, namedRangeSurface);
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
        dev.erst.gridgrind.excel.ExcelRangeSelection.All.class,
        cast(WorkbookReadCommand.GetDataValidations.class, validations).selection());
    assertInstanceOf(
        dev.erst.gridgrind.excel.ExcelRangeSelection.Selected.class,
        cast(WorkbookReadCommand.GetConditionalFormatting.class, conditionalFormatting)
            .selection());
    assertInstanceOf(
        dev.erst.gridgrind.excel.ExcelTableSelection.ByNames.class,
        cast(WorkbookReadCommand.GetTables.class, tables).selection());
    assertInstanceOf(
        dev.erst.gridgrind.excel.ExcelSheetSelection.Selected.class,
        cast(WorkbookReadCommand.GetFormulaSurface.class, formulaSurface).selection());
    assertInstanceOf(
        dev.erst.gridgrind.excel.ExcelNamedRangeSelection.All.class,
        cast(WorkbookReadCommand.GetNamedRangeSurface.class, namedRangeSurface).selection());
  }

  @Test
  void convertsAnalysisReadOperationsIntoWorkbookReadCommands() {
    WorkbookReadCommand formulaHealth =
        WorkbookReadCommandConverter.toReadCommand(
            new WorkbookReadOperation.AnalyzeFormulaHealth(
                "formulaHealth", new SheetSelection.All()));
    WorkbookReadCommand validationHealth =
        WorkbookReadCommandConverter.toReadCommand(
            new WorkbookReadOperation.AnalyzeDataValidationHealth(
                "validationHealth", new SheetSelection.All()));
    WorkbookReadCommand conditionalFormattingHealth =
        WorkbookReadCommandConverter.toReadCommand(
            new WorkbookReadOperation.AnalyzeConditionalFormattingHealth(
                "conditionalFormattingHealth", new SheetSelection.Selected(List.of("Budget"))));
    WorkbookReadCommand autofilterHealth =
        WorkbookReadCommandConverter.toReadCommand(
            new WorkbookReadOperation.AnalyzeAutofilterHealth(
                "autofilterHealth", new SheetSelection.Selected(List.of("Budget"))));
    WorkbookReadCommand tableHealth =
        WorkbookReadCommandConverter.toReadCommand(
            new WorkbookReadOperation.AnalyzeTableHealth(
                "tableHealth", new TableSelection.ByNames(List.of("BudgetTable"))));
    WorkbookReadCommand hyperlinkHealth =
        WorkbookReadCommandConverter.toReadCommand(
            new WorkbookReadOperation.AnalyzeHyperlinkHealth(
                "hyperlinkHealth", new SheetSelection.All()));
    WorkbookReadCommand namedRangeHealth =
        WorkbookReadCommandConverter.toReadCommand(
            new WorkbookReadOperation.AnalyzeNamedRangeHealth(
                "namedRangeHealth", new NamedRangeSelection.All()));
    WorkbookReadCommand workbookFindings =
        WorkbookReadCommandConverter.toReadCommand(
            new WorkbookReadOperation.AnalyzeWorkbookFindings("workbookFindings"));

    assertInstanceOf(WorkbookReadCommand.AnalyzeFormulaHealth.class, formulaHealth);
    assertInstanceOf(WorkbookReadCommand.AnalyzeDataValidationHealth.class, validationHealth);
    assertInstanceOf(
        WorkbookReadCommand.AnalyzeConditionalFormattingHealth.class, conditionalFormattingHealth);
    assertInstanceOf(WorkbookReadCommand.AnalyzeAutofilterHealth.class, autofilterHealth);
    assertInstanceOf(WorkbookReadCommand.AnalyzeTableHealth.class, tableHealth);
    assertInstanceOf(WorkbookReadCommand.AnalyzeHyperlinkHealth.class, hyperlinkHealth);
    assertInstanceOf(WorkbookReadCommand.AnalyzeNamedRangeHealth.class, namedRangeHealth);
    assertInstanceOf(WorkbookReadCommand.AnalyzeWorkbookFindings.class, workbookFindings);
    assertInstanceOf(
        dev.erst.gridgrind.excel.ExcelSheetSelection.Selected.class,
        cast(
                WorkbookReadCommand.AnalyzeConditionalFormattingHealth.class,
                conditionalFormattingHealth)
            .selection());
  }

  @Test
  void convertsSheetStateOperationsIntoWorkbookCommandsWithExactFields() {
    WorkbookCommand.CopySheet copySheet =
        cast(
            WorkbookCommand.CopySheet.class,
            WorkbookCommandConverter.toCommand(
                new WorkbookOperation.CopySheet(
                    "Budget", "Budget Copy", new SheetCopyPosition.AtIndex(1))));
    WorkbookCommand.SetActiveSheet setActiveSheet =
        cast(
            WorkbookCommand.SetActiveSheet.class,
            WorkbookCommandConverter.toCommand(
                new WorkbookOperation.SetActiveSheet("Budget Copy")));
    WorkbookCommand.SetSelectedSheets setSelectedSheets =
        cast(
            WorkbookCommand.SetSelectedSheets.class,
            WorkbookCommandConverter.toCommand(
                new WorkbookOperation.SetSelectedSheets(List.of("Budget", "Budget Copy"))));
    WorkbookCommand.SetSheetVisibility setSheetVisibility =
        cast(
            WorkbookCommand.SetSheetVisibility.class,
            WorkbookCommandConverter.toCommand(
                new WorkbookOperation.SetSheetVisibility("Budget", ExcelSheetVisibility.HIDDEN)));
    WorkbookCommand.SetSheetProtection setSheetProtection =
        cast(
            WorkbookCommand.SetSheetProtection.class,
            WorkbookCommandConverter.toCommand(
                new WorkbookOperation.SetSheetProtection("Budget", protectionSettings())));
    WorkbookCommand.ClearSheetProtection clearSheetProtection =
        cast(
            WorkbookCommand.ClearSheetProtection.class,
            WorkbookCommandConverter.toCommand(
                new WorkbookOperation.ClearSheetProtection("Budget")));

    ExcelSheetCopyPosition.AtIndex position =
        assertInstanceOf(ExcelSheetCopyPosition.AtIndex.class, copySheet.position());

    assertEquals("Budget", copySheet.sourceSheetName());
    assertEquals("Budget Copy", copySheet.newSheetName());
    assertEquals(1, position.targetIndex());
    assertEquals("Budget Copy", setActiveSheet.sheetName());
    assertEquals(List.of("Budget", "Budget Copy"), setSelectedSheets.sheetNames());
    assertEquals(ExcelSheetVisibility.HIDDEN, setSheetVisibility.visibility());
    assertEquals(excelProtectionSettings(), setSheetProtection.protection());
    assertEquals("Budget", clearSheetProtection.sheetName());
  }

  @Test
  void convertsReadResultsIntoProtocolReadResults() {
    ExcelCellSnapshot.BlankSnapshot blank =
        new ExcelCellSnapshot.BlankSnapshot(
            "A1", "BLANK", "", defaultStyle(), ExcelCellMetadataSnapshot.empty());

    WorkbookReadResult workbookSummary =
        WorkbookReadResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummaryResult(
                "workbook",
                new dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummary.WithSheets(
                    1, List.of("Budget"), "Budget", List.of("Budget"), 1, true)));
    WorkbookReadResult namedRanges =
        WorkbookReadResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.NamedRangesResult(
                "ranges",
                List.of(
                    new ExcelNamedRangeSnapshot.RangeSnapshot(
                        "BudgetTotal",
                        new ExcelNamedRangeScope.WorkbookScope(),
                        "Budget!$B$4",
                        new ExcelNamedRangeTarget("Budget", "B4")))));
    WorkbookReadResult sheetSummary =
        WorkbookReadResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.SheetSummaryResult(
                "sheet",
                new dev.erst.gridgrind.excel.WorkbookReadResult.SheetSummary(
                    "Budget",
                    dev.erst.gridgrind.excel.ExcelSheetVisibility.VISIBLE,
                    new dev.erst.gridgrind.excel.WorkbookReadResult.SheetProtection.Unprotected(),
                    4,
                    3,
                    2)));
    WorkbookReadResult cells =
        WorkbookReadResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.CellsResult(
                "cells", "Budget", List.of(blank)));
    WorkbookReadResult window =
        WorkbookReadResultConverter.toReadResult(
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
        WorkbookReadResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.MergedRegionsResult(
                "merged",
                "Budget",
                List.of(new dev.erst.gridgrind.excel.WorkbookReadResult.MergedRegion("A1:B2"))));
    WorkbookReadResult hyperlinks =
        WorkbookReadResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.HyperlinksResult(
                "hyperlinks",
                "Budget",
                List.of(
                    new dev.erst.gridgrind.excel.WorkbookReadResult.CellHyperlink(
                        "A1", new ExcelHyperlink.Url("https://example.com/report")))));
    WorkbookReadResult comments =
        WorkbookReadResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.CommentsResult(
                "comments",
                "Budget",
                List.of(
                    new dev.erst.gridgrind.excel.WorkbookReadResult.CellComment(
                        "A1", new ExcelComment("Review", "GridGrind", false)))));
    WorkbookReadResult layout =
        WorkbookReadResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.SheetLayoutResult(
                "layout",
                new dev.erst.gridgrind.excel.WorkbookReadResult.SheetLayout(
                    "Budget",
                    new dev.erst.gridgrind.excel.ExcelSheetPane.Frozen(1, 1, 1, 1),
                    125,
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookReadResult.ColumnLayout(
                            0, 12.5, false, 0, false)),
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookReadResult.RowLayout(
                            0, 18.0, false, 0, false)))));
    WorkbookReadResult printLayout =
        WorkbookReadResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.PrintLayoutResult(
                "printLayout",
                "Budget",
                new dev.erst.gridgrind.excel.ExcelPrintLayout(
                    new dev.erst.gridgrind.excel.ExcelPrintLayout.Area.Range("A1:B20"),
                    ExcelPrintOrientation.LANDSCAPE,
                    new dev.erst.gridgrind.excel.ExcelPrintLayout.Scaling.Fit(1, 0),
                    new dev.erst.gridgrind.excel.ExcelPrintLayout.TitleRows.Band(0, 0),
                    new dev.erst.gridgrind.excel.ExcelPrintLayout.TitleColumns.Band(0, 0),
                    new dev.erst.gridgrind.excel.ExcelHeaderFooterText("Budget", "", ""),
                    new dev.erst.gridgrind.excel.ExcelHeaderFooterText("", "Page &P", ""))));
    WorkbookReadResult conditionalFormatting =
        WorkbookReadResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.ConditionalFormattingResult(
                "conditional-formatting",
                "Budget",
                List.of(
                    new ExcelConditionalFormattingBlockSnapshot(
                        List.of("A2:A5"),
                        List.of(
                            new ExcelConditionalFormattingRuleSnapshot.FormulaRule(
                                1,
                                true,
                                "A2>0",
                                new ExcelDifferentialStyleSnapshot(
                                    "0.00", true, null, null, "#102030", null, null, "#E0F0AA",
                                    null, List.of())),
                            new ExcelConditionalFormattingRuleSnapshot.ColorScaleRule(
                                2,
                                false,
                                List.of(
                                    new ExcelConditionalFormattingThresholdSnapshot(
                                        ExcelConditionalFormattingThresholdType.MIN, null, null),
                                    new ExcelConditionalFormattingThresholdSnapshot(
                                        ExcelConditionalFormattingThresholdType.MAX, null, null)),
                                List.of("#AA0000", "#00AA00")))))));
    WorkbookReadResult formulaSurface =
        WorkbookReadResultConverter.toReadResult(
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
        WorkbookReadResultConverter.toReadResult(
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
        WorkbookReadResultConverter.toReadResult(
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
    WorkbookReadResult formulaHealth =
        WorkbookReadResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.FormulaHealthResult(
                "formula-health",
                new dev.erst.gridgrind.excel.WorkbookAnalysis.FormulaHealth(
                    1,
                    new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSummary(1, 0, 0, 1),
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFinding(
                            dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFindingCode
                                .FORMULA_VOLATILE_FUNCTION,
                            dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSeverity.INFO,
                            "Volatile formula",
                            "Formula uses NOW().",
                            new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisLocation.Cell(
                                "Budget", "B4"),
                            List.of("NOW()"))))));

    assertInstanceOf(WorkbookReadResult.WorkbookSummaryResult.class, workbookSummary);
    assertInstanceOf(WorkbookReadResult.NamedRangesResult.class, namedRanges);
    assertInstanceOf(WorkbookReadResult.SheetSummaryResult.class, sheetSummary);
    assertInstanceOf(WorkbookReadResult.CellsResult.class, cells);
    assertInstanceOf(WorkbookReadResult.WindowResult.class, window);
    assertInstanceOf(WorkbookReadResult.MergedRegionsResult.class, merged);
    assertInstanceOf(WorkbookReadResult.HyperlinksResult.class, hyperlinks);
    assertInstanceOf(WorkbookReadResult.CommentsResult.class, comments);
    assertInstanceOf(WorkbookReadResult.SheetLayoutResult.class, layout);
    assertInstanceOf(WorkbookReadResult.PrintLayoutResult.class, printLayout);
    assertInstanceOf(WorkbookReadResult.ConditionalFormattingResult.class, conditionalFormatting);
    assertInstanceOf(WorkbookReadResult.FormulaSurfaceResult.class, formulaSurface);
    assertInstanceOf(WorkbookReadResult.SheetSchemaResult.class, schema);
    assertInstanceOf(WorkbookReadResult.NamedRangeSurfaceResult.class, namedRangeSurface);
    assertInstanceOf(WorkbookReadResult.FormulaHealthResult.class, formulaHealth);
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
        new HyperlinkTarget.Url("https://example.com/report"),
        cast(WorkbookReadResult.HyperlinksResult.class, hyperlinks)
            .hyperlinks()
            .getFirst()
            .hyperlink());
    assertEquals(
        "Review",
        cast(WorkbookReadResult.CommentsResult.class, comments)
            .comments()
            .getFirst()
            .comment()
            .text());
    assertInstanceOf(
        PaneReport.Frozen.class,
        cast(WorkbookReadResult.SheetLayoutResult.class, layout).layout().pane());
    assertEquals(
        125, cast(WorkbookReadResult.SheetLayoutResult.class, layout).layout().zoomPercent());
    assertEquals(
        ExcelPrintOrientation.LANDSCAPE,
        cast(WorkbookReadResult.PrintLayoutResult.class, printLayout).layout().orientation());
    assertEquals(
        2,
        cast(WorkbookReadResult.ConditionalFormattingResult.class, conditionalFormatting)
            .conditionalFormattingBlocks()
            .getFirst()
            .rules()
            .size());
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
    assertEquals(
        1,
        cast(WorkbookReadResult.FormulaHealthResult.class, formulaHealth)
            .analysis()
            .summary()
            .infoCount());
  }

  @Test
  void convertsRemainingAnalysisReadResultsIntoProtocolReadResults() {
    WorkbookReadResult conditionalFormattingHealth =
        WorkbookReadResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.ConditionalFormattingHealthResult(
                "conditional-formatting-health",
                new dev.erst.gridgrind.excel.WorkbookAnalysis.ConditionalFormattingHealth(
                    1,
                    new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSummary(1, 1, 0, 0),
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFinding(
                            dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFindingCode
                                .CONDITIONAL_FORMATTING_PRIORITY_COLLISION,
                            dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSeverity.ERROR,
                            "Priority collision",
                            "Conditional-formatting priorities collide.",
                            new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisLocation.Sheet(
                                "Budget"),
                            List.of("FORMULA_RULE@Budget!A1:A3"))))));
    WorkbookReadResult hyperlinkHealth =
        WorkbookReadResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.HyperlinkHealthResult(
                "hyperlink-health",
                new dev.erst.gridgrind.excel.WorkbookAnalysis.HyperlinkHealth(
                    2,
                    new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSummary(2, 1, 1, 0),
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFinding(
                            dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFindingCode
                                .HYPERLINK_MALFORMED_TARGET,
                            dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSeverity.ERROR,
                            "Malformed hyperlink target",
                            "Hyperlink target is malformed.",
                            new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisLocation
                                .Workbook(),
                            List.of("mailto:")),
                        new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFinding(
                            dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFindingCode
                                .HYPERLINK_MISSING_DOCUMENT_SHEET,
                            dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSeverity.WARNING,
                            "Missing target sheet",
                            "Sheet is missing.",
                            new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisLocation.Sheet(
                                "Budget"),
                            List.of("Missing!A1"))))));
    WorkbookReadResult namedRangeHealth =
        WorkbookReadResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.NamedRangeHealthResult(
                "named-range-health",
                new dev.erst.gridgrind.excel.WorkbookAnalysis.NamedRangeHealth(
                    1,
                    new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSummary(1, 0, 1, 0),
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFinding(
                            dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFindingCode
                                .NAMED_RANGE_BROKEN_REFERENCE,
                            dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSeverity.WARNING,
                            "Broken named range",
                            "Named range contains #REF!.",
                            new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisLocation.Range(
                                "Budget", "A1:B2"),
                            List.of("#REF!"))))));
    WorkbookReadResult workbookFindings =
        WorkbookReadResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookFindingsResult(
                "workbook-findings",
                new dev.erst.gridgrind.excel.WorkbookAnalysis.WorkbookFindings(
                    new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSummary(1, 0, 0, 1),
                    List.of(
                        new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFinding(
                            dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisFindingCode
                                .NAMED_RANGE_SCOPE_SHADOWING,
                            dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisSeverity.INFO,
                            "Scope shadowing",
                            "Name exists in more than one scope.",
                            new dev.erst.gridgrind.excel.WorkbookAnalysis.AnalysisLocation
                                .NamedRange(
                                "BudgetTotal", new ExcelNamedRangeScope.SheetScope("Budget")),
                            List.of("Budget!$B$4"))))));

    assertEquals(
        1,
        cast(
                WorkbookReadResult.ConditionalFormattingHealthResult.class,
                conditionalFormattingHealth)
            .analysis()
            .checkedConditionalFormattingBlockCount());
    GridGrindResponse.AnalysisLocationReport workbookLocation =
        cast(WorkbookReadResult.HyperlinkHealthResult.class, hyperlinkHealth)
            .analysis()
            .findings()
            .getFirst()
            .location();
    GridGrindResponse.AnalysisLocationReport sheetLocation =
        cast(WorkbookReadResult.HyperlinkHealthResult.class, hyperlinkHealth)
            .analysis()
            .findings()
            .get(1)
            .location();
    GridGrindResponse.AnalysisLocationReport rangeLocation =
        cast(WorkbookReadResult.NamedRangeHealthResult.class, namedRangeHealth)
            .analysis()
            .findings()
            .getFirst()
            .location();
    GridGrindResponse.AnalysisLocationReport namedRangeLocation =
        cast(WorkbookReadResult.WorkbookFindingsResult.class, workbookFindings)
            .analysis()
            .findings()
            .getFirst()
            .location();

    assertInstanceOf(GridGrindResponse.AnalysisLocationReport.Workbook.class, workbookLocation);
    assertEquals(
        "Budget",
        cast(GridGrindResponse.AnalysisLocationReport.Sheet.class, sheetLocation).sheetName());
    assertEquals(
        "A1:B2", cast(GridGrindResponse.AnalysisLocationReport.Range.class, rangeLocation).range());
    assertEquals(
        "BudgetTotal",
        cast(GridGrindResponse.AnalysisLocationReport.NamedRange.class, namedRangeLocation).name());
  }

  @Test
  void convertsNamedRangeFormulaSnapshotsAndFormulaBackedSurfaceEntries() {
    GridGrindResponse.NamedRangeReport formulaReport =
        WorkbookReadResultConverter.toNamedRangeReport(
            new ExcelNamedRangeSnapshot.FormulaSnapshot(
                "BudgetRollup", new ExcelNamedRangeScope.WorkbookScope(), "SUM(Budget!$B$2:$B$3)"));
    assertInstanceOf(GridGrindResponse.NamedRangeReport.FormulaReport.class, formulaReport);

    WorkbookReadResult.NamedRangeSurfaceResult surface =
        cast(
            WorkbookReadResult.NamedRangeSurfaceResult.class,
            WorkbookReadResultConverter.toReadResult(
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
  void convertsPaneNoneIntoProtocolReport() {
    WorkbookReadResult.SheetLayoutResult layout =
        cast(
            WorkbookReadResult.SheetLayoutResult.class,
            WorkbookReadResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookReadResult.SheetLayoutResult(
                    "layout",
                    new dev.erst.gridgrind.excel.WorkbookReadResult.SheetLayout(
                        "Budget",
                        new dev.erst.gridgrind.excel.ExcelSheetPane.None(),
                        100,
                        List.of(),
                        List.of()))));

    assertInstanceOf(PaneReport.None.class, layout.layout().pane());
  }

  @Test
  void convertsSplitPaneAndDefaultPrintLayoutIntoProtocolReports() {
    WorkbookReadResult.SheetLayoutResult layout =
        cast(
            WorkbookReadResult.SheetLayoutResult.class,
            WorkbookReadResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookReadResult.SheetLayoutResult(
                    "layout",
                    new dev.erst.gridgrind.excel.WorkbookReadResult.SheetLayout(
                        "Budget",
                        new dev.erst.gridgrind.excel.ExcelSheetPane.Split(
                            1200, 2400, 3, 4, ExcelPaneRegion.LOWER_RIGHT),
                        100,
                        List.of(),
                        List.of()))));
    WorkbookReadResult.PrintLayoutResult printLayout =
        cast(
            WorkbookReadResult.PrintLayoutResult.class,
            WorkbookReadResultConverter.toReadResult(
                new dev.erst.gridgrind.excel.WorkbookReadResult.PrintLayoutResult(
                    "print-layout",
                    "Budget",
                    new dev.erst.gridgrind.excel.ExcelPrintLayout(
                        new dev.erst.gridgrind.excel.ExcelPrintLayout.Area.None(),
                        ExcelPrintOrientation.PORTRAIT,
                        new dev.erst.gridgrind.excel.ExcelPrintLayout.Scaling.Automatic(),
                        new dev.erst.gridgrind.excel.ExcelPrintLayout.TitleRows.None(),
                        new dev.erst.gridgrind.excel.ExcelPrintLayout.TitleColumns.None(),
                        new dev.erst.gridgrind.excel.ExcelHeaderFooterText("", "", ""),
                        new dev.erst.gridgrind.excel.ExcelHeaderFooterText("", "", "")))));

    PaneReport.Split pane = assertInstanceOf(PaneReport.Split.class, layout.layout().pane());
    assertEquals(1200, pane.xSplitPosition());
    assertEquals(2400, pane.ySplitPosition());
    assertEquals(3, pane.leftmostColumn());
    assertEquals(4, pane.topRow());
    assertEquals(ExcelPaneRegion.LOWER_RIGHT, pane.activePane());
    assertInstanceOf(PrintAreaReport.None.class, printLayout.layout().printArea());
    assertInstanceOf(PrintScalingReport.Automatic.class, printLayout.layout().scaling());
    assertInstanceOf(PrintTitleRowsReport.None.class, printLayout.layout().repeatingRows());
    assertInstanceOf(PrintTitleColumnsReport.None.class, printLayout.layout().repeatingColumns());
  }

  @Test
  void convertsSheetStateReadResultsIntoProtocolShapes() {
    WorkbookReadResult emptyWorkbookSummary =
        WorkbookReadResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummaryResult(
                "workbook",
                new dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummary.Empty(
                    0, List.of(), 0, false)));
    WorkbookReadResult populatedWorkbookSummary =
        WorkbookReadResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummaryResult(
                "workbook-2",
                new dev.erst.gridgrind.excel.WorkbookReadResult.WorkbookSummary.WithSheets(
                    2,
                    List.of("Budget", "Archive"),
                    "Archive",
                    List.of("Budget", "Archive"),
                    1,
                    true)));
    WorkbookReadResult protectedSheetSummary =
        WorkbookReadResultConverter.toReadResult(
            new dev.erst.gridgrind.excel.WorkbookReadResult.SheetSummaryResult(
                "sheet",
                new dev.erst.gridgrind.excel.WorkbookReadResult.SheetSummary(
                    "Budget",
                    dev.erst.gridgrind.excel.ExcelSheetVisibility.VERY_HIDDEN,
                    new dev.erst.gridgrind.excel.WorkbookReadResult.SheetProtection.Protected(
                        excelProtectionSettings()),
                    4,
                    7,
                    3)));

    GridGrindResponse.WorkbookSummary.Empty empty =
        assertInstanceOf(
            GridGrindResponse.WorkbookSummary.Empty.class,
            cast(WorkbookReadResult.WorkbookSummaryResult.class, emptyWorkbookSummary).workbook());
    GridGrindResponse.WorkbookSummary.WithSheets populated =
        assertInstanceOf(
            GridGrindResponse.WorkbookSummary.WithSheets.class,
            cast(WorkbookReadResult.WorkbookSummaryResult.class, populatedWorkbookSummary)
                .workbook());
    GridGrindResponse.SheetSummaryReport sheet =
        cast(WorkbookReadResult.SheetSummaryResult.class, protectedSheetSummary).sheet();
    GridGrindResponse.SheetProtectionReport.Protected protection =
        assertInstanceOf(
            GridGrindResponse.SheetProtectionReport.Protected.class, sheet.protection());

    assertEquals(0, empty.sheetCount());
    assertEquals("Archive", populated.activeSheetName());
    assertEquals(List.of("Budget", "Archive"), populated.selectedSheetNames());
    assertEquals(ExcelSheetVisibility.VERY_HIDDEN, sheet.visibility());
    assertEquals(protectionSettings(), protection.settings());
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
            "GET_PRINT_LAYOUT",
            "GET_DATA_VALIDATIONS",
            "GET_CONDITIONAL_FORMATTING",
            "GET_AUTOFILTERS",
            "GET_TABLES",
            "GET_FORMULA_SURFACE",
            "GET_SHEET_SCHEMA",
            "GET_NAMED_RANGE_SURFACE",
            "ANALYZE_FORMULA_HEALTH",
            "ANALYZE_DATA_VALIDATION_HEALTH",
            "ANALYZE_CONDITIONAL_FORMATTING_HEALTH",
            "ANALYZE_AUTOFILTER_HEALTH",
            "ANALYZE_TABLE_HEALTH",
            "ANALYZE_HYPERLINK_HEALTH",
            "ANALYZE_NAMED_RANGE_HEALTH",
            "ANALYZE_WORKBOOK_FINDINGS"),
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
                new WorkbookReadOperation.GetPrintLayout("print-layout", "Budget")),
            DefaultGridGrindRequestExecutor.readType(
                new WorkbookReadOperation.GetDataValidations(
                    "validations", "Budget", new RangeSelection.All())),
            DefaultGridGrindRequestExecutor.readType(
                new WorkbookReadOperation.GetConditionalFormatting(
                    "conditional-formatting", "Budget", new RangeSelection.All())),
            DefaultGridGrindRequestExecutor.readType(
                new WorkbookReadOperation.GetAutofilters("autofilters", "Budget")),
            DefaultGridGrindRequestExecutor.readType(
                new WorkbookReadOperation.GetTables(
                    "tables", new TableSelection.ByNames(List.of("BudgetTable")))),
            DefaultGridGrindRequestExecutor.readType(
                new WorkbookReadOperation.GetFormulaSurface("formula", new SheetSelection.All())),
            DefaultGridGrindRequestExecutor.readType(
                new WorkbookReadOperation.GetSheetSchema("schema", "Budget", "A1", 1, 1)),
            DefaultGridGrindRequestExecutor.readType(
                new WorkbookReadOperation.GetNamedRangeSurface(
                    "surface", new NamedRangeSelection.All())),
            DefaultGridGrindRequestExecutor.readType(
                new WorkbookReadOperation.AnalyzeFormulaHealth(
                    "formula-health", new SheetSelection.All())),
            DefaultGridGrindRequestExecutor.readType(
                new WorkbookReadOperation.AnalyzeDataValidationHealth(
                    "validation-health", new SheetSelection.All())),
            DefaultGridGrindRequestExecutor.readType(
                new WorkbookReadOperation.AnalyzeConditionalFormattingHealth(
                    "conditional-formatting-health", new SheetSelection.All())),
            DefaultGridGrindRequestExecutor.readType(
                new WorkbookReadOperation.AnalyzeAutofilterHealth(
                    "autofilter-health", new SheetSelection.All())),
            DefaultGridGrindRequestExecutor.readType(
                new WorkbookReadOperation.AnalyzeTableHealth(
                    "table-health", new TableSelection.All())),
            DefaultGridGrindRequestExecutor.readType(
                new WorkbookReadOperation.AnalyzeHyperlinkHealth(
                    "hyperlink-health", new SheetSelection.All())),
            DefaultGridGrindRequestExecutor.readType(
                new WorkbookReadOperation.AnalyzeNamedRangeHealth(
                    "named-range-health", new NamedRangeSelection.All())),
            DefaultGridGrindRequestExecutor.readType(
                new WorkbookReadOperation.AnalyzeWorkbookFindings("workbook-findings"))));
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
    WorkbookReadOperation printLayout =
        new WorkbookReadOperation.GetPrintLayout("print-layout", "Budget");
    WorkbookReadOperation validations =
        new WorkbookReadOperation.GetDataValidations(
            "validations", "Budget", new RangeSelection.Selected(List.of("A1:A3")));
    WorkbookReadOperation conditionalFormatting =
        new WorkbookReadOperation.GetConditionalFormatting(
            "conditional-formatting", "Budget", new RangeSelection.Selected(List.of("B2:B5")));
    WorkbookReadOperation autofilters =
        new WorkbookReadOperation.GetAutofilters("autofilters", "Budget");
    WorkbookReadOperation tables =
        new WorkbookReadOperation.GetTables(
            "tables", new TableSelection.ByNames(List.of("BudgetTable")));
    WorkbookReadOperation formula =
        new WorkbookReadOperation.GetFormulaSurface("formula", new SheetSelection.All());
    WorkbookReadOperation schema =
        new WorkbookReadOperation.GetSheetSchema("schema", "Budget", "C3", 2, 2);
    WorkbookReadOperation surface =
        new WorkbookReadOperation.GetNamedRangeSurface("surface", new NamedRangeSelection.All());
    WorkbookReadOperation formulaHealth =
        new WorkbookReadOperation.AnalyzeFormulaHealth(
            "formula-health", new SheetSelection.Selected(List.of("Budget")));
    WorkbookReadOperation validationHealth =
        new WorkbookReadOperation.AnalyzeDataValidationHealth(
            "validation-health", new SheetSelection.Selected(List.of("Budget")));
    WorkbookReadOperation conditionalFormattingHealth =
        new WorkbookReadOperation.AnalyzeConditionalFormattingHealth(
            "conditional-formatting-health", new SheetSelection.Selected(List.of("Budget")));
    WorkbookReadOperation autofilterHealth =
        new WorkbookReadOperation.AnalyzeAutofilterHealth(
            "autofilter-health", new SheetSelection.Selected(List.of("Budget")));
    WorkbookReadOperation tableHealth =
        new WorkbookReadOperation.AnalyzeTableHealth(
            "table-health", new TableSelection.ByNames(List.of("BudgetTable")));
    WorkbookReadOperation hyperlinkHealth =
        new WorkbookReadOperation.AnalyzeHyperlinkHealth(
            "hyperlink-health", new SheetSelection.All());
    WorkbookReadOperation namedRangeHealth =
        new WorkbookReadOperation.AnalyzeNamedRangeHealth(
            "named-range-health",
            new NamedRangeSelection.Selected(
                List.of(new NamedRangeSelector.WorkbookScope("BudgetTotal"))));
    WorkbookReadOperation workbookFindings =
        new WorkbookReadOperation.AnalyzeWorkbookFindings("workbook-findings");

    assertReadContext(workbook, null, null, null, runtimeException);
    assertReadContext(namedRanges, null, null, null, runtimeException);
    assertReadContext(sheet, "Budget", null, null, runtimeException);
    assertReadContext(cells, "Budget", null, null, runtimeException);
    assertReadContext(window, "Budget", "B2", null, runtimeException);
    assertReadContext(merged, "Budget", null, null, runtimeException);
    assertReadContext(hyperlinks, "Budget", null, null, runtimeException);
    assertReadContext(comments, "Budget", null, null, runtimeException);
    assertReadContext(layout, "Budget", null, null, runtimeException);
    assertReadContext(printLayout, "Budget", null, null, runtimeException);
    assertReadContext(validations, "Budget", null, null, runtimeException);
    assertReadContext(conditionalFormatting, "Budget", null, null, runtimeException);
    assertReadContext(autofilters, "Budget", null, null, runtimeException);
    assertReadContext(tables, null, null, null, runtimeException);
    assertReadContext(formula, null, null, null, runtimeException);
    assertReadContext(schema, "Budget", "C3", null, runtimeException);
    assertReadContext(surface, null, null, null, runtimeException);
    assertReadContext(formulaHealth, "Budget", null, null, runtimeException);
    assertReadContext(validationHealth, "Budget", null, null, runtimeException);
    assertReadContext(conditionalFormattingHealth, "Budget", null, null, runtimeException);
    assertReadContext(autofilterHealth, "Budget", null, null, runtimeException);
    assertReadContext(tableHealth, null, null, null, runtimeException);
    assertReadContext(hyperlinkHealth, null, null, null, runtimeException);
    assertReadContext(namedRangeHealth, null, null, "BudgetTotal", runtimeException);
    assertReadContext(workbookFindings, null, null, null, runtimeException);

    assertEquals(
        "BudgetTotal",
        DefaultGridGrindRequestExecutor.namedRangeNameFor(workbook, missingNamedRange));
    assertEquals(
        "BudgetTotal",
        DefaultGridGrindRequestExecutor.namedRangeNameFor(namedRanges, missingNamedRange));
    assertEquals("BAD!", DefaultGridGrindRequestExecutor.addressFor(cells, invalidAddress));
  }

  @Test
  void extractsSingleSheetAndNamedRangeContextOnlyWhenSelectionsAreUnambiguous() {
    assertEquals(
        "Budget",
        DefaultGridGrindRequestExecutor.sheetNameFor(
            new WorkbookReadOperation.GetFormulaSurface(
                "formula", new SheetSelection.Selected(List.of("Budget")))));
    assertNull(
        DefaultGridGrindRequestExecutor.sheetNameFor(
            new WorkbookReadOperation.AnalyzeHyperlinkHealth(
                "hyperlink-health", new SheetSelection.Selected(List.of("Budget", "Forecast")))));

    assertEquals(
        "BudgetTotal",
        DefaultGridGrindRequestExecutor.namedRangeNameFor(
            new WorkbookReadOperation.GetNamedRangeSurface(
                "surface",
                new NamedRangeSelection.Selected(
                    List.of(new NamedRangeSelector.ByName("BudgetTotal")))),
            new RuntimeException("x")));
    assertEquals(
        "BudgetTotal",
        DefaultGridGrindRequestExecutor.namedRangeNameFor(
            new WorkbookReadOperation.GetNamedRangeSurface(
                "surface",
                new NamedRangeSelection.Selected(
                    List.of(new NamedRangeSelector.SheetScope("BudgetTotal", "Budget")))),
            new RuntimeException("x")));
    assertNull(
        DefaultGridGrindRequestExecutor.namedRangeNameFor(
            new WorkbookReadOperation.GetNamedRangeSurface(
                "surface",
                new NamedRangeSelection.Selected(
                    List.of(
                        new NamedRangeSelector.WorkbookScope("BudgetTotal"),
                        new NamedRangeSelector.WorkbookScope("ForecastTotal")))),
            new RuntimeException("x")));
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
            new WorkbookReadOperation.GetSheetSchema("schema", "Budget", "C3", 2, 2),
            invalidAddress));
    assertEquals(
        "BudgetTotal",
        DefaultGridGrindRequestExecutor.namedRangeNameFor(
            new WorkbookReadOperation.GetNamedRangeSurface(
                "surface", new NamedRangeSelection.All()),
            missingNamedRange));
  }

  @Test
  void extractsContextForAuthoringMetadataAndNamedRangeOperations() {
    RuntimeException exception = new RuntimeException("test");
    assertWriteContext(
        new WorkbookOperation.SetHyperlink(
            "Budget", "A1", new HyperlinkTarget.Url("https://example.com/report")),
        exception,
        "Budget",
        "A1",
        null,
        null);
    assertWriteContext(
        new WorkbookOperation.ClearHyperlink("Budget", "A1"),
        exception,
        "Budget",
        "A1",
        null,
        null);
    assertWriteContext(
        new WorkbookOperation.SetComment(
            "Budget", "A1", new CommentInput("Review", "GridGrind", false)),
        exception,
        "Budget",
        "A1",
        null,
        null);
    assertWriteContext(
        new WorkbookOperation.ClearComment("Budget", "A1"), exception, "Budget", "A1", null, null);
    assertWriteContext(
        new WorkbookOperation.ApplyStyle(
            "Budget",
            "A1:B2",
            new CellStyleInput(
                null,
                null,
                new CellFontInput(true, null, null, null, null, null, null),
                null,
                null,
                null)),
        exception,
        "Budget",
        null,
        "A1:B2",
        null);
    assertWriteContext(
        new WorkbookOperation.SetDataValidation(
            "Budget",
            "B2:B5",
            new DataValidationInput(
                new DataValidationRuleInput.TextLength(
                    ExcelComparisonOperator.LESS_OR_EQUAL, "20", null),
                true,
                false,
                new DataValidationPromptInput("Reason", "Use 20 characters or fewer.", true),
                null)),
        exception,
        "Budget",
        null,
        "B2:B5",
        null);
    assertWriteContext(
        new WorkbookOperation.ClearDataValidations("Budget", new RangeSelection.All()),
        exception,
        "Budget",
        null,
        null,
        null);
    assertWriteContext(
        new WorkbookOperation.SetConditionalFormatting(
            "Budget",
            new ConditionalFormattingBlockInput(
                List.of("B2:B5"),
                List.of(
                    new ConditionalFormattingRuleInput.FormulaRule(
                        "B2>0",
                        true,
                        new DifferentialStyleInput(
                            null, true, null, null, null, null, null, null, null))))),
        exception,
        "Budget",
        null,
        "B2:B5",
        null);
    assertWriteContext(
        new WorkbookOperation.ClearConditionalFormatting("Budget", new RangeSelection.All()),
        exception,
        "Budget",
        null,
        null,
        null);
    assertWriteContext(
        new WorkbookOperation.SetAutofilter("Budget", "A1:C4"),
        exception,
        "Budget",
        null,
        "A1:C4",
        null);
    assertWriteContext(
        new WorkbookOperation.ClearAutofilter("Budget"), exception, "Budget", null, null, null);
    assertWriteContext(
        new WorkbookOperation.SetTable(
            new TableInput("BudgetTable", "Budget", "A1:C4", false, new TableStyleInput.None())),
        exception,
        "Budget",
        null,
        "A1:C4",
        null);
    assertWriteContext(
        new WorkbookOperation.DeleteTable("BudgetTable", "Budget"),
        exception,
        "Budget",
        null,
        null,
        null);
    assertWriteContext(
        new WorkbookOperation.SetNamedRange(
            "BudgetTotal", new NamedRangeScope.Workbook(), new NamedRangeTarget("Budget", "B4")),
        exception,
        "Budget",
        null,
        "B4",
        "BudgetTotal");
    assertWriteContext(
        new WorkbookOperation.DeleteNamedRange("BudgetTotal", new NamedRangeScope.Workbook()),
        exception,
        null,
        null,
        null,
        "BudgetTotal");
    assertWriteContext(
        new WorkbookOperation.DeleteNamedRange("LocalItem", new NamedRangeScope.Sheet("Budget")),
        exception,
        "Budget",
        null,
        null,
        "LocalItem");
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
    assertEquals(
        "B2:B5",
        DefaultGridGrindRequestExecutor.rangeFor(
            new WorkbookOperation.SetConditionalFormatting(
                "Budget",
                new ConditionalFormattingBlockInput(
                    List.of("B2:B5"),
                    List.of(
                        new ConditionalFormattingRuleInput.FormulaRule(
                            "B2>0",
                            true,
                            new DifferentialStyleInput(
                                null, true, null, null, null, null, null, null, null))))),
            invalidFormula));
    assertNull(
        DefaultGridGrindRequestExecutor.rangeFor(
            new WorkbookOperation.SetConditionalFormatting(
                "Budget",
                new ConditionalFormattingBlockInput(
                    List.of("B2:B5", "D2:D5"),
                    List.of(
                        new ConditionalFormattingRuleInput.FormulaRule(
                            "B2>0",
                            true,
                            new DifferentialStyleInput(
                                null, true, null, null, null, null, null, null, null))))),
            invalidFormula));
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
            new WorkbookOperation.InsertRows("Budget", 1, 2),
            new WorkbookOperation.DeleteRows("Budget", new RowSpanInput.Band(1, 2)),
            new WorkbookOperation.ShiftRows("Budget", new RowSpanInput.Band(1, 2), 1),
            new WorkbookOperation.InsertColumns("Budget", 1, 2),
            new WorkbookOperation.DeleteColumns("Budget", new ColumnSpanInput.Band(1, 2)),
            new WorkbookOperation.ShiftColumns("Budget", new ColumnSpanInput.Band(1, 2), -1),
            new WorkbookOperation.SetRowVisibility("Budget", new RowSpanInput.Band(1, 2), true),
            new WorkbookOperation.SetColumnVisibility(
                "Budget", new ColumnSpanInput.Band(1, 2), false),
            new WorkbookOperation.GroupRows("Budget", new RowSpanInput.Band(1, 2), true),
            new WorkbookOperation.UngroupRows("Budget", new RowSpanInput.Band(1, 2)),
            new WorkbookOperation.GroupColumns("Budget", new ColumnSpanInput.Band(1, 2), true),
            new WorkbookOperation.UngroupColumns("Budget", new ColumnSpanInput.Band(1, 2)),
            new WorkbookOperation.SetSheetPane("Budget", new PaneInput.Frozen(1, 1, 1, 1)),
            new WorkbookOperation.SetSheetZoom("Budget", 125),
            new WorkbookOperation.SetPrintLayout(
                "Budget",
                new PrintLayoutInput(
                    new PrintAreaInput.Range("A1:B12"),
                    ExcelPrintOrientation.LANDSCAPE,
                    new PrintScalingInput.Fit(1, 0),
                    new PrintTitleRowsInput.Band(0, 0),
                    new PrintTitleColumnsInput.Band(0, 0),
                    new HeaderFooterTextInput("Budget", "", ""),
                    new HeaderFooterTextInput("", "Page &P", ""))),
            new WorkbookOperation.ClearPrintLayout("Budget"),
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
                    null,
                    null,
                    new CellFontInput(true, null, null, null, null, null, null),
                    null,
                    null,
                    null)),
            new WorkbookOperation.SetDataValidation(
                "Budget",
                "B2:B5",
                new DataValidationInput(
                    new DataValidationRuleInput.TextLength(
                        ExcelComparisonOperator.LESS_OR_EQUAL, "20", null),
                    true,
                    false,
                    null,
                    null)),
            new WorkbookOperation.ClearDataValidations("Budget", new RangeSelection.All()),
            new WorkbookOperation.SetConditionalFormatting(
                "Budget",
                new ConditionalFormattingBlockInput(
                    List.of("B2:B5"),
                    List.of(
                        new ConditionalFormattingRuleInput.FormulaRule(
                            "B2>0",
                            true,
                            new DifferentialStyleInput(
                                null, true, null, null, null, null, null, null, null))))),
            new WorkbookOperation.ClearConditionalFormatting("Budget", new RangeSelection.All()),
            new WorkbookOperation.SetAutofilter("Budget", "A1:C4"),
            new WorkbookOperation.ClearAutofilter("Budget"),
            new WorkbookOperation.SetTable(
                new TableInput(
                    "BudgetTable", "Budget", "A1:C4", false, new TableStyleInput.None())),
            new WorkbookOperation.DeleteTable("BudgetTable", "Budget"),
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
            new WorkbookOperation.SetSheetPane("Budget", new PaneInput.Frozen(1, 1, 1, 1)),
            missingNamedRange));
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
    return switch (success.persistence()) {
      case GridGrindResponse.PersistenceOutcome.SavedAs savedAs -> savedAs.executionPath();
      case GridGrindResponse.PersistenceOutcome.Overwritten overwritten ->
          overwritten.executionPath();
      case GridGrindResponse.PersistenceOutcome.NotSaved _ ->
          throw new AssertionError("expected persisted workbook");
    };
  }

  private static List<String> requestIds(GridGrindResponse.Success success) {
    return success.reads().stream().map(WorkbookReadResult::requestId).toList();
  }

  private static <T> T cast(Class<T> type, Object value) {
    return type.cast(assertInstanceOf(type, value));
  }

  private static void assertReadContext(
      WorkbookReadOperation operation,
      String expectedSheetName,
      String expectedRuntimeAddress,
      String expectedNamedRangeName,
      RuntimeException runtimeException) {
    assertEquals(expectedSheetName, DefaultGridGrindRequestExecutor.sheetNameFor(operation));
    assertEquals(
        expectedRuntimeAddress,
        DefaultGridGrindRequestExecutor.addressFor(operation, runtimeException));
    assertEquals(
        expectedNamedRangeName,
        DefaultGridGrindRequestExecutor.namedRangeNameFor(operation, runtimeException));
  }

  private static void assertWriteContext(
      WorkbookOperation operation,
      Exception exception,
      String expectedSheetName,
      String expectedAddress,
      String expectedRange,
      String expectedNamedRangeName) {
    assertNull(DefaultGridGrindRequestExecutor.formulaFor(operation, exception));
    assertEquals(
        expectedSheetName, DefaultGridGrindRequestExecutor.sheetNameFor(operation, exception));
    assertEquals(expectedAddress, DefaultGridGrindRequestExecutor.addressFor(operation, exception));
    assertEquals(expectedRange, DefaultGridGrindRequestExecutor.rangeFor(operation, exception));
    assertEquals(
        expectedNamedRangeName,
        DefaultGridGrindRequestExecutor.namedRangeNameFor(operation, exception));
  }

  @Test
  void workbookLocationForUsesExistingSourcePathsWhenPersistenceDoesNotSave() {
    Path existingWorkbookPath = Path.of("tmp", "existing-budget.xlsx").toAbsolutePath();

    WorkbookLocation workbookLocation =
        DefaultGridGrindRequestExecutor.workbookLocationFor(
            new GridGrindRequest.WorkbookSource.ExistingFile(existingWorkbookPath.toString()),
            new GridGrindRequest.WorkbookPersistence.None());

    WorkbookLocation.StoredWorkbook storedWorkbook =
        assertInstanceOf(WorkbookLocation.StoredWorkbook.class, workbookLocation);
    assertEquals(existingWorkbookPath.normalize(), storedWorkbook.workbookPath());
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
    return WorkbookReadResultConverter.toCellStyleReport(style);
  }

  private static ExcelCellStyleSnapshot defaultStyle() {
    return new ExcelCellStyleSnapshot(
        "",
        new ExcelCellAlignmentSnapshot(
            false, ExcelHorizontalAlignment.GENERAL, ExcelVerticalAlignment.BOTTOM, 0, 0),
        new ExcelCellFontSnapshot(
            false,
            false,
            "Aptos",
            ExcelFontHeight.fromPoints(new BigDecimal("11")),
            null,
            false,
            false),
        new ExcelCellFillSnapshot(dev.erst.gridgrind.excel.ExcelFillPattern.NONE, null, null),
        new ExcelBorderSnapshot(
            new ExcelBorderSide(ExcelBorderStyle.NONE, null),
            new ExcelBorderSide(ExcelBorderStyle.NONE, null),
            new ExcelBorderSide(ExcelBorderStyle.NONE, null),
            new ExcelBorderSide(ExcelBorderStyle.NONE, null)),
        new ExcelCellProtectionSnapshot(true, false));
  }

  private static SheetProtectionSettings protectionSettings() {
    return new SheetProtectionSettings(
        false, true, false, true, false, true, false, true, false, true, false, true, false, true,
        false);
  }

  private static ExcelSheetProtectionSettings excelProtectionSettings() {
    return new ExcelSheetProtectionSettings(
        false, true, false, true, false, true, false, true, false, true, false, true, false, true,
        false);
  }
}
