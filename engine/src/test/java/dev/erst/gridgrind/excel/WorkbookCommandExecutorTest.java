package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Integration tests for WorkbookCommandExecutor applying commands to a workbook. */
class WorkbookCommandExecutorTest {
  @Test
  void appliesAllSupportedCommandTypesThroughVarargs() throws IOException {
    var workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-command-layout-");

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      WorkbookCommandExecutor executor = new WorkbookCommandExecutor();

      assertSame(
          workbook,
          executor.apply(
              workbook,
              new WorkbookCommand.CreateSheet("Budget"),
              new WorkbookCommand.CreateSheet("Archive"),
              new WorkbookCommand.CreateSheet("Scratch"),
              new WorkbookCommand.CreateSheet("Ops"),
              new WorkbookCommand.SetRange(
                  "Budget",
                  "A1:B2",
                  List.of(
                      List.of(ExcelCellValue.text("Item"), ExcelCellValue.number(49.0)),
                      List.of(ExcelCellValue.text("Tax"), ExcelCellValue.number(5.0)))),
              new WorkbookCommand.MergeCells("Budget", "A1:B1"),
              new WorkbookCommand.ApplyStyle(
                  "Budget",
                  "A1:B1",
                  ExcelCellStyle.alignment(
                      ExcelHorizontalAlignment.CENTER, ExcelVerticalAlignment.CENTER)),
              new WorkbookCommand.AppendRow(
                  "Budget",
                  List.of(ExcelCellValue.text("Total"), ExcelCellValue.formula("SUM(B1:B2)"))),
              new WorkbookCommand.ClearRange("Budget", "A2"),
              new WorkbookCommand.SetHyperlink(
                  "Budget", "A1", new ExcelHyperlink.Document("Budget!B3")),
              new WorkbookCommand.ClearHyperlink("Budget", "A1"),
              new WorkbookCommand.SetComment(
                  "Budget", "B1", new ExcelComment("Review", "GridGrind", false)),
              new WorkbookCommand.ClearComment("Budget", "B1"),
              new WorkbookCommand.SetDataValidation("Budget", "C1:C3", validationDefinition()),
              new WorkbookCommand.ClearDataValidations(
                  "Budget", new ExcelRangeSelection.Selected(List.of("C2"))),
              new WorkbookCommand.SetAutofilter("Budget", "A1:B3"),
              new WorkbookCommand.ClearAutofilter("Budget"),
              new WorkbookCommand.SetRange(
                  "Ops",
                  "A1:B3",
                  List.of(
                      List.of(ExcelCellValue.text("Owner"), ExcelCellValue.text("Task")),
                      List.of(ExcelCellValue.text("Ada"), ExcelCellValue.text("Queue")),
                      List.of(ExcelCellValue.text("Lin"), ExcelCellValue.text("Pack")))),
              new WorkbookCommand.SetRange(
                  "Ops",
                  "D1:E3",
                  List.of(
                      List.of(ExcelCellValue.text("Region"), ExcelCellValue.text("Desk")),
                      List.of(ExcelCellValue.text("North"), ExcelCellValue.text("A1")),
                      List.of(ExcelCellValue.text("South"), ExcelCellValue.text("B1")))),
              new WorkbookCommand.SetRange(
                  "Ops",
                  "G1:H3",
                  List.of(
                      List.of(ExcelCellValue.text("Stage"), ExcelCellValue.text("Team")),
                      List.of(ExcelCellValue.text("Review"), ExcelCellValue.text("Docs")),
                      List.of(ExcelCellValue.text("Ship"), ExcelCellValue.text("Ops")))),
              new WorkbookCommand.SetAutofilter("Ops", "D1:E3"),
              new WorkbookCommand.SetTable(
                  new ExcelTableDefinition(
                      "Queue", "Ops", "A1:B3", false, new ExcelTableStyle.None())),
              new WorkbookCommand.SetTable(
                  new ExcelTableDefinition(
                      "Stages",
                      "Ops",
                      "G1:H3",
                      false,
                      new ExcelTableStyle.Named("TableStyleMedium2", false, false, true, false))),
              new WorkbookCommand.DeleteTable("Stages", "Ops"),
              new WorkbookCommand.RenameSheet("Archive", "History"),
              new WorkbookCommand.MoveSheet("History", 0),
              new WorkbookCommand.DeleteSheet("Scratch"),
              new WorkbookCommand.SetNamedRange(
                  new ExcelNamedRangeDefinition(
                      "BudgetTotal",
                      new ExcelNamedRangeScope.WorkbookScope(),
                      new ExcelNamedRangeTarget("Budget", "B3"))),
              new WorkbookCommand.SetNamedRange(
                  new ExcelNamedRangeDefinition(
                      "HistoryCell",
                      new ExcelNamedRangeScope.SheetScope("History"),
                      new ExcelNamedRangeTarget("History", "A1"))),
              new WorkbookCommand.DeleteNamedRange(
                  "BudgetTotal", new ExcelNamedRangeScope.WorkbookScope()),
              new WorkbookCommand.AutoSizeColumns("Budget"),
              new WorkbookCommand.SetColumnWidth("Budget", 0, 1, 16.0),
              new WorkbookCommand.SetRowHeight("Budget", 0, 0, 28.5),
              new WorkbookCommand.SetSheetPane("Budget", new ExcelSheetPane.Frozen(1, 1, 1, 1)),
              new WorkbookCommand.SetSheetZoom("Budget", 125),
              new WorkbookCommand.SetPrintLayout(
                  "Budget",
                  new ExcelPrintLayout(
                      new ExcelPrintLayout.Area.Range("A1:B3"),
                      ExcelPrintOrientation.LANDSCAPE,
                      new ExcelPrintLayout.Scaling.Fit(1, 0),
                      new ExcelPrintLayout.TitleRows.Band(0, 0),
                      new ExcelPrintLayout.TitleColumns.Band(0, 0),
                      new ExcelHeaderFooterText("Budget", "", ""),
                      new ExcelHeaderFooterText("", "Confidential", ""))),
              new WorkbookCommand.UnmergeCells("Budget", "A1:B1"),
              new WorkbookCommand.EvaluateAllFormulas(),
              new WorkbookCommand.ForceFormulaRecalculationOnOpen()));
      assertEquals(List.of("History", "Budget", "Ops"), workbook.sheetNames());
      assertEquals("Item", workbook.sheet("Budget").text("A1"));
      assertEquals(54.0, workbook.sheet("Budget").number("B3"));
      assertTrue(workbook.sheet("Budget").snapshotCell("A1").metadata().hyperlink().isEmpty());
      assertTrue(workbook.sheet("Budget").snapshotCell("B1").metadata().comment().isEmpty());
      assertEquals(1, workbook.namedRangeCount());
      assertEquals(
          ExcelHorizontalAlignment.CENTER,
          workbook.sheet("Budget").snapshotCell("A1").style().alignment().horizontalAlignment());
      assertEquals("BLANK", workbook.sheet("Budget").snapshotCell("A2").effectiveType());
      assertThrows(SheetNotFoundException.class, () -> workbook.sheet("Scratch"));
      assertTrue(workbook.forceFormulaRecalculationOnOpenEnabled());
      WorkbookReadResult.AutofiltersResult autofilters =
          assertInstanceOf(
              WorkbookReadResult.AutofiltersResult.class,
              new ExcelWorkbookIntrospector()
                  .execute(workbook, new WorkbookReadCommand.GetAutofilters("autofilters", "Ops")));
      WorkbookReadResult.TablesResult tables =
          assertInstanceOf(
              WorkbookReadResult.TablesResult.class,
              new ExcelWorkbookIntrospector()
                  .execute(
                      workbook,
                      new WorkbookReadCommand.GetTables("tables", new ExcelTableSelection.All())));
      assertEquals(
          List.of(
              new ExcelAutofilterSnapshot.SheetOwned("D1:E3"),
              new ExcelAutofilterSnapshot.TableOwned("A1:B3", "Queue")),
          autofilters.autofilters());
      assertEquals(
          List.of("Queue"), tables.tables().stream().map(ExcelTableSnapshot::name).toList());
      workbook.save(workbookPath);
    }

    assertEquals(List.of(), XlsxRoundTrip.mergedRegions(workbookPath, "Budget"));
    assertEquals(4096, XlsxRoundTrip.columnWidth(workbookPath, "Budget", 0));
    assertEquals((short) 570, XlsxRoundTrip.rowHeightTwips(workbookPath, "Budget", 0));
    assertEquals(new ExcelSheetPane.Frozen(1, 1, 1, 1), XlsxRoundTrip.pane(workbookPath, "Budget"));
    assertEquals(125, XlsxRoundTrip.zoomPercent(workbookPath, "Budget"));
    assertEquals(
        ExcelPrintOrientation.LANDSCAPE,
        XlsxRoundTrip.printLayout(workbookPath, "Budget").orientation());
    assertEquals(
        List.of(
            new ExcelDataValidationSnapshot.Supported(List.of("C1", "C3"), validationDefinition())),
        XlsxRoundTrip.dataValidations(workbookPath, "Budget"));
    assertEquals(
        List.of(
            new ExcelNamedRangeSnapshot.RangeSnapshot(
                "HistoryCell",
                new ExcelNamedRangeScope.SheetScope("History"),
                "History!$A$1",
                new ExcelNamedRangeTarget("History", "A1"))),
        XlsxRoundTrip.namedRanges(workbookPath));
    try (ExcelWorkbook reopened = ExcelWorkbook.open(workbookPath)) {
      WorkbookReadResult.AutofiltersResult autofilters =
          assertInstanceOf(
              WorkbookReadResult.AutofiltersResult.class,
              new ExcelWorkbookIntrospector()
                  .execute(reopened, new WorkbookReadCommand.GetAutofilters("autofilters", "Ops")));
      WorkbookReadResult.TablesResult tables =
          assertInstanceOf(
              WorkbookReadResult.TablesResult.class,
              new ExcelWorkbookIntrospector()
                  .execute(
                      reopened,
                      new WorkbookReadCommand.GetTables("tables", new ExcelTableSelection.All())));
      assertEquals(
          List.of(
              new ExcelAutofilterSnapshot.SheetOwned("D1:E3"),
              new ExcelAutofilterSnapshot.TableOwned("A1:B3", "Queue")),
          autofilters.autofilters());
      assertEquals(1, tables.tables().size());
      assertEquals("Queue", tables.tables().getFirst().name());
      assertTrue(tables.tables().getFirst().hasAutofilter());
    }
  }

  @Test
  void rejectsMutationCommandsWhenSheetDoesNotExist() throws IOException {
    WorkbookCommandExecutor executor = new WorkbookCommandExecutor();

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      assertThrows(
          SheetNotFoundException.class,
          () ->
              executor.apply(
                  workbook,
                  new WorkbookCommand.SetCell("Missing", "A1", ExcelCellValue.text("x"))));
      assertThrows(
          SheetNotFoundException.class,
          () ->
              executor.apply(
                  workbook,
                  new WorkbookCommand.SetRange(
                      "Missing", "A1", List.of(List.of(ExcelCellValue.text("x"))))));
      assertThrows(
          SheetNotFoundException.class,
          () ->
              executor.apply(
                  workbook,
                  new WorkbookCommand.SetHyperlink(
                      "Missing", "A1", new ExcelHyperlink.Document("Budget!A1"))));
      assertThrows(
          SheetNotFoundException.class,
          () ->
              executor.apply(
                  workbook,
                  new WorkbookCommand.SetComment(
                      "Missing", "A1", new ExcelComment("Review", "GridGrind", false))));
      assertThrows(
          SheetNotFoundException.class,
          () ->
              executor.apply(
                  workbook,
                  new WorkbookCommand.ApplyStyle(
                      "Missing",
                      "A1",
                      ExcelCellStyle.alignment(
                          ExcelHorizontalAlignment.CENTER, ExcelVerticalAlignment.TOP))));
      assertThrows(
          SheetNotFoundException.class,
          () ->
              executor.apply(
                  workbook,
                  new WorkbookCommand.SetDataValidation("Missing", "A1", validationDefinition())));
      assertThrows(
          SheetNotFoundException.class,
          () ->
              executor.apply(
                  workbook,
                  new WorkbookCommand.SetConditionalFormatting(
                      "Missing",
                      new ExcelConditionalFormattingBlockDefinition(
                          List.of("A1"),
                          List.of(
                              new ExcelConditionalFormattingRule.FormulaRule(
                                  "A1>0",
                                  false,
                                  new ExcelDifferentialStyle(
                                      null, true, null, null, null, null, null, null, null)))))));
      assertThrows(
          SheetNotFoundException.class,
          () ->
              executor.apply(
                  workbook,
                  new WorkbookCommand.ClearDataValidations(
                      "Missing", new ExcelRangeSelection.All())));
      assertThrows(
          SheetNotFoundException.class,
          () ->
              executor.apply(
                  workbook,
                  new WorkbookCommand.ClearConditionalFormatting(
                      "Missing", new ExcelRangeSelection.All())));
      assertThrows(
          SheetNotFoundException.class,
          () -> executor.apply(workbook, new WorkbookCommand.SetAutofilter("Missing", "A1:B2")));
      assertThrows(
          SheetNotFoundException.class,
          () -> executor.apply(workbook, new WorkbookCommand.ClearAutofilter("Missing")));
      assertThrows(
          SheetNotFoundException.class,
          () ->
              executor.apply(
                  workbook,
                  new WorkbookCommand.SetTable(
                      new ExcelTableDefinition(
                          "Queue", "Missing", "A1:B2", false, new ExcelTableStyle.None()))));
      assertThrows(
          IllegalArgumentException.class,
          () -> executor.apply(workbook, new WorkbookCommand.DeleteTable("Queue", "Missing")));
      assertThrows(
          SheetNotFoundException.class,
          () ->
              executor.apply(
                  workbook,
                  new WorkbookCommand.AppendRow("Missing", List.of(ExcelCellValue.text("x")))));
    }
  }

  @Test
  void appliesDrawingCommandsThroughTheWorkbookCommandSurface() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      WorkbookCommandExecutor executor = new WorkbookCommandExecutor();
      ExcelDrawingAnchor.TwoCell firstAnchor =
          new ExcelDrawingAnchor.TwoCell(
              new ExcelDrawingMarker(1, 1),
              new ExcelDrawingMarker(4, 6),
              ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
      ExcelDrawingAnchor.TwoCell movedAnchor =
          new ExcelDrawingAnchor.TwoCell(
              new ExcelDrawingMarker(6, 2),
              new ExcelDrawingMarker(9, 7),
              ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE);
      byte[] pngBytes =
          java.util.Base64.getDecoder()
              .decode(
                  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2kQAAAAASUVORK5CYII=");

      executor.apply(
          workbook,
          new WorkbookCommand.CreateSheet("Ops"),
          new WorkbookCommand.SetPicture(
              "Ops",
              new ExcelPictureDefinition(
                  "OpsPicture",
                  new ExcelBinaryData(pngBytes),
                  ExcelPictureFormat.PNG,
                  firstAnchor,
                  "Queue preview")),
          new WorkbookCommand.SetShape(
              "Ops",
              new ExcelShapeDefinition(
                  "OpsShape",
                  ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE,
                  firstAnchor,
                  "rect",
                  "Queue")),
          new WorkbookCommand.SetEmbeddedObject(
              "Ops",
              new ExcelEmbeddedObjectDefinition(
                  "OpsEmbed",
                  "Payload",
                  "payload.txt",
                  "payload.txt",
                  new ExcelBinaryData("payload".getBytes(java.nio.charset.StandardCharsets.UTF_8)),
                  ExcelPictureFormat.PNG,
                  new ExcelBinaryData(pngBytes),
                  firstAnchor)),
          new WorkbookCommand.SetDrawingObjectAnchor("Ops", "OpsShape", movedAnchor),
          new WorkbookCommand.DeleteDrawingObject("Ops", "OpsPicture"));

      List<ExcelDrawingObjectSnapshot> drawingObjects = workbook.sheet("Ops").drawingObjects();
      assertEquals(
          List.of("OpsShape", "OpsEmbed"),
          drawingObjects.stream().map(ExcelDrawingObjectSnapshot::name).toList());
      assertEquals(
          movedAnchor,
          assertInstanceOf(ExcelDrawingObjectSnapshot.Shape.class, drawingObjects.getFirst())
              .anchor());
    }
  }

  @Test
  void appliesWorkbookProtectionCommands() throws IOException {
    WorkbookCommandExecutor executor = new WorkbookCommandExecutor();

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Budget");

      executor.apply(
          workbook,
          new WorkbookCommand.SetWorkbookProtection(
              new ExcelWorkbookProtectionSettings(true, false, true, "secret", "review")));

      assertEquals(
          new ExcelWorkbookProtectionSnapshot(true, false, true, true, true),
          workbook.workbookProtection());

      executor.apply(workbook, new WorkbookCommand.ClearWorkbookProtection());

      assertEquals(
          new ExcelWorkbookProtectionSnapshot(false, false, false, false, false),
          workbook.workbookProtection());
    }
  }

  @Test
  void appliesClearPrintLayoutThroughIterableCommands() throws IOException {
    WorkbookCommandExecutor executor = new WorkbookCommandExecutor();

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Budget");
      workbook
          .sheet("Budget")
          .setPrintLayout(
              new ExcelPrintLayout(
                  new ExcelPrintLayout.Area.Range("A1:B3"),
                  ExcelPrintOrientation.LANDSCAPE,
                  new ExcelPrintLayout.Scaling.Fit(1, 0),
                  new ExcelPrintLayout.TitleRows.Band(0, 0),
                  new ExcelPrintLayout.TitleColumns.Band(0, 0),
                  new ExcelHeaderFooterText("Budget", "", ""),
                  new ExcelHeaderFooterText("", "Confidential", "")));

      assertSame(
          workbook,
          executor.apply(workbook, List.of(new WorkbookCommand.ClearPrintLayout("Budget"))));
      ExcelPrintLayout clearedPrintLayout = workbook.sheet("Budget").printLayout();
      assertEquals(new ExcelPrintLayout.Area.None(), clearedPrintLayout.printArea());
      assertEquals(ExcelPrintOrientation.PORTRAIT, clearedPrintLayout.orientation());
      assertEquals(new ExcelPrintLayout.Scaling.Automatic(), clearedPrintLayout.scaling());
      assertEquals(new ExcelPrintLayout.TitleRows.None(), clearedPrintLayout.repeatingRows());
      assertEquals(new ExcelPrintLayout.TitleColumns.None(), clearedPrintLayout.repeatingColumns());
      assertEquals(ExcelHeaderFooterText.blank(), clearedPrintLayout.header());
      assertEquals(ExcelHeaderFooterText.blank(), clearedPrintLayout.footer());
    }
  }

  @Test
  void appliesConditionalFormattingCommandsThroughExecutor() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      WorkbookCommandExecutor executor = new WorkbookCommandExecutor();
      workbook.getOrCreateSheet("Ops");

      executor.apply(
          workbook,
          new WorkbookCommand.SetConditionalFormatting(
              "Ops",
              new ExcelConditionalFormattingBlockDefinition(
                  List.of("A1:A3"),
                  List.of(
                      new ExcelConditionalFormattingRule.FormulaRule(
                          "A1>0",
                          true,
                          new ExcelDifferentialStyle(
                              null, true, null, null, "#112233", null, null, null, null))))));

      WorkbookReadResult.ConditionalFormattingResult initial =
          assertInstanceOf(
              WorkbookReadResult.ConditionalFormattingResult.class,
              new ExcelWorkbookIntrospector()
                  .execute(
                      workbook,
                      new WorkbookReadCommand.GetConditionalFormatting(
                          "cf", "Ops", new ExcelRangeSelection.All())));
      assertEquals(1, initial.conditionalFormattingBlocks().size());

      executor.apply(
          workbook,
          new WorkbookCommand.ClearConditionalFormatting("Ops", new ExcelRangeSelection.All()));

      WorkbookReadResult.ConditionalFormattingResult cleared =
          assertInstanceOf(
              WorkbookReadResult.ConditionalFormattingResult.class,
              new ExcelWorkbookIntrospector()
                  .execute(
                      workbook,
                      new WorkbookReadCommand.GetConditionalFormatting(
                          "cf", "Ops", new ExcelRangeSelection.All())));
      assertEquals(List.of(), cleared.conditionalFormattingBlocks());
    }
  }

  @Test
  void appliesRowAndColumnStructuralCommandsThroughExecutor() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-structural-commands-");

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      WorkbookCommandExecutor executor = new WorkbookCommandExecutor();

      assertSame(
          workbook,
          executor.apply(
              workbook,
              new WorkbookCommand.CreateSheet("Layout"),
              new WorkbookCommand.SetRange(
                  "Layout",
                  "A1:F6",
                  List.of(
                      List.of(
                          ExcelCellValue.text("Item"),
                          ExcelCellValue.text("Qty"),
                          ExcelCellValue.text("Status"),
                          ExcelCellValue.text("Note"),
                          ExcelCellValue.text("Owner"),
                          ExcelCellValue.text("Flag")),
                      List.of(
                          ExcelCellValue.text("Hosting"),
                          ExcelCellValue.number(42.0),
                          ExcelCellValue.text("Open"),
                          ExcelCellValue.text("Alpha"),
                          ExcelCellValue.text("Ada"),
                          ExcelCellValue.text("Y")),
                      List.of(
                          ExcelCellValue.text("Support"),
                          ExcelCellValue.number(84.0),
                          ExcelCellValue.text("Closed"),
                          ExcelCellValue.text("Beta"),
                          ExcelCellValue.text("Lin"),
                          ExcelCellValue.text("N")),
                      List.of(
                          ExcelCellValue.text("Ops"),
                          ExcelCellValue.number(168.0),
                          ExcelCellValue.text("Open"),
                          ExcelCellValue.text("Gamma"),
                          ExcelCellValue.text("Bea"),
                          ExcelCellValue.text("Y")),
                      List.of(
                          ExcelCellValue.text("QA"),
                          ExcelCellValue.number(21.0),
                          ExcelCellValue.text("Queued"),
                          ExcelCellValue.text("Delta"),
                          ExcelCellValue.text("Kai"),
                          ExcelCellValue.text("N")),
                      List.of(
                          ExcelCellValue.text("Infra"),
                          ExcelCellValue.number(7.0),
                          ExcelCellValue.text("Done"),
                          ExcelCellValue.text("Epsilon"),
                          ExcelCellValue.text("Mia"),
                          ExcelCellValue.text("Y")))),
              new WorkbookCommand.GroupRows("Layout", new ExcelRowSpan(1, 3), true),
              new WorkbookCommand.SetRowVisibility("Layout", new ExcelRowSpan(5, 5), true),
              new WorkbookCommand.GroupColumns("Layout", new ExcelColumnSpan(1, 3), true),
              new WorkbookCommand.SetColumnVisibility("Layout", new ExcelColumnSpan(5, 5), true),
              new WorkbookCommand.CreateSheet("Moves"),
              new WorkbookCommand.SetRange(
                  "Moves",
                  "A1:D3",
                  List.of(
                      List.of(
                          ExcelCellValue.text("Item"),
                          ExcelCellValue.text("Qty"),
                          ExcelCellValue.text("Status"),
                          ExcelCellValue.text("Note")),
                      List.of(
                          ExcelCellValue.text("Hosting"),
                          ExcelCellValue.number(42.0),
                          ExcelCellValue.text("Open"),
                          ExcelCellValue.text("Alpha")),
                      List.of(
                          ExcelCellValue.text("Support"),
                          ExcelCellValue.number(84.0),
                          ExcelCellValue.text("Closed"),
                          ExcelCellValue.text("Beta")))),
              new WorkbookCommand.InsertRows("Moves", 1, 1),
              new WorkbookCommand.SetCell("Moves", "A2", ExcelCellValue.text("Spacer")),
              new WorkbookCommand.ShiftRows("Moves", new ExcelRowSpan(2, 3), 1),
              new WorkbookCommand.DeleteRows("Moves", new ExcelRowSpan(2, 2)),
              new WorkbookCommand.InsertColumns("Moves", 1, 1),
              new WorkbookCommand.SetCell("Moves", "B1", ExcelCellValue.text("Pad")),
              new WorkbookCommand.ShiftColumns("Moves", new ExcelColumnSpan(2, 4), 1),
              new WorkbookCommand.DeleteColumns("Moves", new ExcelColumnSpan(2, 2))));

      assertEquals("Pad", workbook.sheet("Moves").text("B1"));
      assertEquals("Hosting", workbook.sheet("Moves").text("A3"));
      assertEquals(42.0, workbook.sheet("Moves").number("C3"));
      assertEquals("Beta", workbook.sheet("Moves").text("E4"));
      workbook.save(workbookPath);
    }

    WorkbookReadResult.SheetLayout layout = XlsxRoundTrip.sheetLayout(workbookPath, "Layout");
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

    try (ExcelWorkbook reopened = ExcelWorkbook.open(workbookPath)) {
      assertEquals("Pad", reopened.sheet("Moves").text("B1"));
      assertEquals("Hosting", reopened.sheet("Moves").text("A3"));
      assertEquals(42.0, reopened.sheet("Moves").number("C3"));
      assertEquals("Beta", reopened.sheet("Moves").text("E4"));
    }
  }

  @Test
  void appliesSheetManagementCommandsThroughExecutor() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      WorkbookCommandExecutor executor = new WorkbookCommandExecutor();

      assertSame(
          workbook,
          executor.apply(
              workbook,
              new WorkbookCommand.CreateSheet("Alpha"),
              new WorkbookCommand.CreateSheet("Beta"),
              new WorkbookCommand.SetCell("Alpha", "A1", ExcelCellValue.text("Live")),
              new WorkbookCommand.CopySheet(
                  "Alpha", "Alpha Copy", new ExcelSheetCopyPosition.AtIndex(1)),
              new WorkbookCommand.SetActiveSheet("Alpha"),
              new WorkbookCommand.SetSelectedSheets(List.of("Alpha", "Beta")),
              new WorkbookCommand.SetSheetVisibility("Beta", ExcelSheetVisibility.HIDDEN),
              new WorkbookCommand.SetSheetProtection("Alpha", protectionSettings()),
              new WorkbookCommand.ClearSheetProtection("Alpha")));

      assertEquals(List.of("Alpha", "Alpha Copy", "Beta"), workbook.sheetNames());
      assertEquals("Live", workbook.sheet("Alpha Copy").text("A1"));

      WorkbookReadResult.WorkbookSummary.WithSheets summary =
          assertInstanceOf(
              WorkbookReadResult.WorkbookSummary.WithSheets.class, workbook.workbookSummary());
      assertEquals("Alpha", summary.activeSheetName());
      assertEquals(List.of("Alpha"), summary.selectedSheetNames());
      assertEquals(ExcelSheetVisibility.HIDDEN, workbook.sheetSummary("Beta").visibility());
      assertInstanceOf(
          WorkbookReadResult.SheetProtection.Unprotected.class,
          workbook.sheetSummary("Alpha").protection());
    }
  }

  @Test
  void appliesUngroupCommandsThroughExecutor() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-command-ungroup-");

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      WorkbookCommandExecutor executor = new WorkbookCommandExecutor();

      executor.apply(
          workbook,
          new WorkbookCommand.CreateSheet("Layout"),
          new WorkbookCommand.SetRange(
              "Layout",
              "A1:D4",
              List.of(
                  List.of(
                      ExcelCellValue.text("Item"),
                      ExcelCellValue.text("Qty"),
                      ExcelCellValue.text("Status"),
                      ExcelCellValue.text("Owner")),
                  List.of(
                      ExcelCellValue.text("Hosting"),
                      ExcelCellValue.number(42.0),
                      ExcelCellValue.text("Open"),
                      ExcelCellValue.text("Ada")),
                  List.of(
                      ExcelCellValue.text("Support"),
                      ExcelCellValue.number(84.0),
                      ExcelCellValue.text("Closed"),
                      ExcelCellValue.text("Lin")),
                  List.of(
                      ExcelCellValue.text("Ops"),
                      ExcelCellValue.number(7.0),
                      ExcelCellValue.text("Done"),
                      ExcelCellValue.text("Mia")))),
          new WorkbookCommand.GroupRows("Layout", new ExcelRowSpan(1, 3), true),
          new WorkbookCommand.GroupColumns("Layout", new ExcelColumnSpan(1, 3), true),
          new WorkbookCommand.UngroupRows("Layout", new ExcelRowSpan(1, 3)),
          new WorkbookCommand.UngroupColumns("Layout", new ExcelColumnSpan(1, 3)));

      workbook.save(workbookPath);
    }

    WorkbookReadResult.SheetLayout layout = XlsxRoundTrip.sheetLayout(workbookPath, "Layout");

    assertTrue(layout.rows().size() >= 4);
    assertFalse(layout.rows().get(1).hidden());
    assertEquals(0, layout.rows().get(1).outlineLevel());
    assertTrue(layout.columns().size() >= 4);
    assertFalse(layout.columns().get(1).hidden());
    assertEquals(0, layout.columns().get(1).outlineLevel());
  }

  @Test
  void appliesFormulaLifecycleCommandsThroughExecutor() throws IOException {
    WorkbookCommandExecutor executor = new WorkbookCommandExecutor();

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Budget");
      workbook.sheet("Budget").setCell("A1", ExcelCellValue.number(2.0d));
      workbook.sheet("Budget").setCell("B1", ExcelCellValue.formula("A1*2"));
      workbook.sheet("Budget").setCell("C1", ExcelCellValue.formula("A1*3"));

      executor.apply(workbook, new WorkbookCommand.EvaluateAllFormulas());
      workbook.sheet("Budget").setCell("A1", ExcelCellValue.number(4.0d));

      executor.apply(
          workbook,
          new WorkbookCommand.EvaluateFormulaCells(
              List.of(new ExcelFormulaCellTarget("Budget", "B1"))));

      assertEquals(
          8.0d,
          ((ExcelCellSnapshot.NumberSnapshot)
                  ((ExcelCellSnapshot.FormulaSnapshot) workbook.sheet("Budget").snapshotCell("B1"))
                      .evaluation())
              .numberValue());
      assertEquals("6.0", cachedFormulaValue(workbook, "Budget", 0, 2));

      executor.apply(workbook, new WorkbookCommand.ClearFormulaCaches());
      assertNull(cachedFormulaValue(workbook, "Budget", 0, 1));
      assertNull(cachedFormulaValue(workbook, "Budget", 0, 2));

      executor.apply(workbook, new WorkbookCommand.ForceFormulaRecalculationOnOpen());
      assertTrue(workbook.forceFormulaRecalculationOnOpenEnabled());
    }
  }

  @Test
  void rejectsUnexpectedCommandsInPrivateCommandFamilies() throws Exception {
    WorkbookCommand unexpected = new WorkbookCommand.EvaluateAllFormulas();

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Budget");

      assertPrivateHelperRejects("applyWorkbookScopeCommand", workbook, unexpected);
      assertPrivateHelperRejects("applySheetStructureCommand", workbook, unexpected);
      assertPrivateHelperRejects("applyCellValueCommand", workbook, unexpected);
      assertPrivateHelperRejects("applyWorkbookMetadataCommand", workbook, unexpected);
      assertPrivateHelperRejects(
          "applyFormulaCommand", workbook, new WorkbookCommand.CreateSheet("Ops"));
    }
  }

  @Test
  void validatesNullWorkbooksCommandsAndCommandEntries() throws IOException {
    WorkbookCommandExecutor executor = new WorkbookCommandExecutor();

    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      assertThrows(
          NullPointerException.class, () -> executor.apply(workbook, (WorkbookCommand[]) null));
      assertThrows(NullPointerException.class, () -> executor.apply(null, List.of()));
      assertThrows(
          NullPointerException.class,
          () -> executor.apply(workbook, (Iterable<WorkbookCommand>) null));
      assertThrows(
          NullPointerException.class,
          () -> executor.apply(workbook, Arrays.asList((WorkbookCommand) null)));
    }
  }

  private static ExcelDataValidationDefinition validationDefinition() {
    return new ExcelDataValidationDefinition(
        new ExcelDataValidationRule.ExplicitList(List.of("Queued", "Done")),
        true,
        false,
        new ExcelDataValidationPrompt("Status", "Pick one workflow state.", true),
        new ExcelDataValidationErrorAlert(
            ExcelDataValidationErrorStyle.STOP,
            "Invalid status",
            "Use one of the allowed values.",
            true));
  }

  private static ExcelSheetProtectionSettings protectionSettings() {
    return new ExcelSheetProtectionSettings(
        true, false, true, false, true, false, true, false, true, false, true, false, true, false,
        true);
  }

  private static void assertPrivateHelperRejects(
      String methodName, ExcelWorkbook workbook, WorkbookCommand command)
      throws ReflectiveOperationException {
    Method method =
        accessibleMethod(
            WorkbookCommandExecutor.class, methodName, ExcelWorkbook.class, WorkbookCommand.class);

    InvocationTargetException failure =
        assertThrows(InvocationTargetException.class, () -> method.invoke(null, workbook, command));
    assertInstanceOf(IllegalStateException.class, failure.getCause());
  }

  @SuppressWarnings("PMD.AvoidAccessibilityAlteration")
  private static Method accessibleMethod(Class<?> type, String name, Class<?>... parameterTypes)
      throws ReflectiveOperationException {
    Method method = type.getDeclaredMethod(name, parameterTypes);
    method.setAccessible(true);
    return method;
  }

  private static String cachedFormulaValue(
      ExcelWorkbook workbook, String sheetName, int rowIndex, int columnIndex) {
    org.apache.poi.xssf.usermodel.XSSFCell cell =
        workbook.xssfWorkbook().getSheet(sheetName).getRow(rowIndex).getCell(columnIndex);
    return cell.getCTCell().isSetV() ? cell.getCTCell().getV() : null;
  }
}
