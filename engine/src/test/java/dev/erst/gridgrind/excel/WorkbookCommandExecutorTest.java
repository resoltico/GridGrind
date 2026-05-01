package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.foundation.ExcelAuthoredDrawingShapeKind;
import dev.erst.gridgrind.excel.foundation.ExcelColumnSpan;
import dev.erst.gridgrind.excel.foundation.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.excel.foundation.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.excel.foundation.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.foundation.ExcelIgnoredErrorType;
import dev.erst.gridgrind.excel.foundation.ExcelPictureFormat;
import dev.erst.gridgrind.excel.foundation.ExcelPivotDataConsolidateFunction;
import dev.erst.gridgrind.excel.foundation.ExcelPrintOrientation;
import dev.erst.gridgrind.excel.foundation.ExcelRowSpan;
import dev.erst.gridgrind.excel.foundation.ExcelSheetVisibility;
import dev.erst.gridgrind.excel.foundation.ExcelVerticalAlignment;
import java.io.IOException;
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
              new WorkbookSheetCommand.CreateSheet("Budget"),
              new WorkbookSheetCommand.CreateSheet("Archive"),
              new WorkbookSheetCommand.CreateSheet("Scratch"),
              new WorkbookSheetCommand.CreateSheet("Ops"),
              new WorkbookCellCommand.SetRange(
                  "Budget",
                  "A1:B2",
                  List.of(
                      List.of(ExcelCellValue.text("Item"), ExcelCellValue.number(49.0)),
                      List.of(ExcelCellValue.text("Tax"), ExcelCellValue.number(5.0)))),
              new WorkbookStructureCommand.MergeCells("Budget", "A1:B1"),
              new WorkbookFormattingCommand.ApplyStyle(
                  "Budget",
                  "A1:B1",
                  ExcelCellStyle.alignment(
                      ExcelHorizontalAlignment.CENTER, ExcelVerticalAlignment.CENTER)),
              new WorkbookCellCommand.AppendRow(
                  "Budget",
                  List.of(ExcelCellValue.text("Total"), ExcelCellValue.formula("SUM(B1:B2)"))),
              new WorkbookCellCommand.ClearRange("Budget", "A2"),
              new WorkbookAnnotationCommand.SetHyperlink(
                  "Budget", "A1", new ExcelHyperlink.Document("Budget!B3")),
              new WorkbookAnnotationCommand.ClearHyperlink("Budget", "A1"),
              new WorkbookAnnotationCommand.SetComment(
                  "Budget", "B1", new ExcelComment("Review", "GridGrind", false)),
              new WorkbookAnnotationCommand.ClearComment("Budget", "B1"),
              new WorkbookFormattingCommand.SetDataValidation(
                  "Budget", "C1:C3", validationDefinition()),
              new WorkbookFormattingCommand.ClearDataValidations(
                  "Budget", new ExcelRangeSelection.Selected(List.of("C2"))),
              new WorkbookTabularCommand.SetAutofilter("Budget", "A1:B3"),
              new WorkbookTabularCommand.ClearAutofilter("Budget"),
              new WorkbookCellCommand.SetRange(
                  "Ops",
                  "A1:B3",
                  List.of(
                      List.of(ExcelCellValue.text("Owner"), ExcelCellValue.text("Task")),
                      List.of(ExcelCellValue.text("Ada"), ExcelCellValue.text("Queue")),
                      List.of(ExcelCellValue.text("Lin"), ExcelCellValue.text("Pack")))),
              new WorkbookCellCommand.SetRange(
                  "Ops",
                  "D1:E3",
                  List.of(
                      List.of(ExcelCellValue.text("Region"), ExcelCellValue.text("Desk")),
                      List.of(ExcelCellValue.text("North"), ExcelCellValue.text("A1")),
                      List.of(ExcelCellValue.text("South"), ExcelCellValue.text("B1")))),
              new WorkbookCellCommand.SetRange(
                  "Ops",
                  "G1:H3",
                  List.of(
                      List.of(ExcelCellValue.text("Stage"), ExcelCellValue.text("Team")),
                      List.of(ExcelCellValue.text("Review"), ExcelCellValue.text("Docs")),
                      List.of(ExcelCellValue.text("Ship"), ExcelCellValue.text("Ops")))),
              new WorkbookTabularCommand.SetAutofilter("Ops", "D1:E3"),
              new WorkbookTabularCommand.SetTable(
                  new ExcelTableDefinition(
                      "Queue", "Ops", "A1:B3", false, new ExcelTableStyle.None())),
              new WorkbookTabularCommand.SetTable(
                  new ExcelTableDefinition(
                      "Stages",
                      "Ops",
                      "G1:H3",
                      false,
                      new ExcelTableStyle.Named("TableStyleMedium2", false, false, true, false))),
              new WorkbookTabularCommand.DeleteTable("Stages", "Ops"),
              new WorkbookSheetCommand.RenameSheet("Archive", "History"),
              new WorkbookSheetCommand.MoveSheet("History", 0),
              new WorkbookSheetCommand.DeleteSheet("Scratch"),
              new WorkbookMetadataCommand.SetNamedRange(
                  new ExcelNamedRangeDefinition(
                      "BudgetTotal",
                      new ExcelNamedRangeScope.WorkbookScope(),
                      new ExcelNamedRangeTarget("Budget", "B3"))),
              new WorkbookMetadataCommand.SetNamedRange(
                  new ExcelNamedRangeDefinition(
                      "HistoryCell",
                      new ExcelNamedRangeScope.SheetScope("History"),
                      new ExcelNamedRangeTarget("History", "A1"))),
              new WorkbookMetadataCommand.DeleteNamedRange(
                  "BudgetTotal", new ExcelNamedRangeScope.WorkbookScope()),
              new WorkbookLayoutCommand.AutoSizeColumns("Budget"),
              new WorkbookStructureCommand.SetColumnWidth("Budget", 0, 1, 16.0),
              new WorkbookStructureCommand.SetRowHeight("Budget", 0, 0, 28.5),
              new WorkbookLayoutCommand.SetSheetPane(
                  "Budget", new ExcelSheetPane.Frozen(1, 1, 1, 1)),
              new WorkbookLayoutCommand.SetSheetZoom("Budget", 125),
              new WorkbookLayoutCommand.SetPrintLayout(
                  "Budget",
                  new ExcelPrintLayout(
                      new ExcelPrintLayout.Area.Range("A1:B3"),
                      ExcelPrintOrientation.LANDSCAPE,
                      new ExcelPrintLayout.Scaling.Fit(1, 0),
                      new ExcelPrintLayout.TitleRows.Band(0, 0),
                      new ExcelPrintLayout.TitleColumns.Band(0, 0),
                      new ExcelHeaderFooterText("Budget", "", ""),
                      new ExcelHeaderFooterText("", "Confidential", ""))),
              new WorkbookStructureCommand.UnmergeCells("Budget", "A1:B1")));
      workbook.formulas().markRecalculateOnOpen();
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
      assertTrue(workbook.formulas().recalculateOnOpenEnabled());
      WorkbookRuleResult.AutofiltersResult autofilters =
          assertInstanceOf(
              WorkbookRuleResult.AutofiltersResult.class,
              new ExcelWorkbookIntrospector()
                  .execute(workbook, new WorkbookReadCommand.GetAutofilters("autofilters", "Ops")));
      WorkbookRuleResult.TablesResult tables =
          assertInstanceOf(
              WorkbookRuleResult.TablesResult.class,
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
      WorkbookRuleResult.AutofiltersResult autofilters =
          assertInstanceOf(
              WorkbookRuleResult.AutofiltersResult.class,
              new ExcelWorkbookIntrospector()
                  .execute(reopened, new WorkbookReadCommand.GetAutofilters("autofilters", "Ops")));
      WorkbookRuleResult.TablesResult tables =
          assertInstanceOf(
              WorkbookRuleResult.TablesResult.class,
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
                  new WorkbookCellCommand.SetCell("Missing", "A1", ExcelCellValue.text("x"))));
      assertThrows(
          SheetNotFoundException.class,
          () ->
              executor.apply(
                  workbook,
                  new WorkbookCellCommand.SetRange(
                      "Missing", "A1", List.of(List.of(ExcelCellValue.text("x"))))));
      assertThrows(
          SheetNotFoundException.class,
          () ->
              executor.apply(
                  workbook,
                  new WorkbookAnnotationCommand.SetHyperlink(
                      "Missing", "A1", new ExcelHyperlink.Document("Budget!A1"))));
      assertThrows(
          SheetNotFoundException.class,
          () ->
              executor.apply(
                  workbook,
                  new WorkbookAnnotationCommand.SetComment(
                      "Missing", "A1", new ExcelComment("Review", "GridGrind", false))));
      assertThrows(
          SheetNotFoundException.class,
          () ->
              executor.apply(
                  workbook,
                  new WorkbookFormattingCommand.ApplyStyle(
                      "Missing",
                      "A1",
                      ExcelCellStyle.alignment(
                          ExcelHorizontalAlignment.CENTER, ExcelVerticalAlignment.TOP))));
      assertThrows(
          SheetNotFoundException.class,
          () ->
              executor.apply(
                  workbook,
                  new WorkbookFormattingCommand.SetDataValidation(
                      "Missing", "A1", validationDefinition())));
      assertThrows(
          SheetNotFoundException.class,
          () ->
              executor.apply(
                  workbook,
                  new WorkbookFormattingCommand.SetConditionalFormatting(
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
                  new WorkbookFormattingCommand.ClearDataValidations(
                      "Missing", new ExcelRangeSelection.All())));
      assertThrows(
          SheetNotFoundException.class,
          () ->
              executor.apply(
                  workbook,
                  new WorkbookFormattingCommand.ClearConditionalFormatting(
                      "Missing", new ExcelRangeSelection.All())));
      assertThrows(
          SheetNotFoundException.class,
          () ->
              executor.apply(
                  workbook, new WorkbookTabularCommand.SetAutofilter("Missing", "A1:B2")));
      assertThrows(
          SheetNotFoundException.class,
          () -> executor.apply(workbook, new WorkbookTabularCommand.ClearAutofilter("Missing")));
      assertThrows(
          SheetNotFoundException.class,
          () ->
              executor.apply(
                  workbook,
                  new WorkbookTabularCommand.SetTable(
                      new ExcelTableDefinition(
                          "Queue", "Missing", "A1:B2", false, new ExcelTableStyle.None()))));
      assertThrows(
          IllegalArgumentException.class,
          () ->
              executor.apply(workbook, new WorkbookTabularCommand.DeleteTable("Queue", "Missing")));
      assertThrows(
          SheetNotFoundException.class,
          () ->
              executor.apply(
                  workbook,
                  new WorkbookCellCommand.AppendRow("Missing", List.of(ExcelCellValue.text("x")))));
    }
  }

  @Test
  void appliesSheetPresentationThroughSheetStructureDispatch() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      WorkbookCommandExecutor executor = new WorkbookCommandExecutor();
      workbook.getOrCreateSheet("Budget");
      ExcelSheetPresentation presentation =
          new ExcelSheetPresentation(
              new ExcelSheetDisplay(false, false, false, true, true),
              ExcelColor.rgb("#225577"),
              new ExcelSheetOutlineSummary(false, false),
              new ExcelSheetDefaults(12, 19.5d),
              List.of(
                  new ExcelIgnoredError(
                      "B2:A1", List.of(ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT))));

      assertSame(
          workbook,
          executor.apply(
              workbook, new WorkbookLayoutCommand.SetSheetPresentation("Budget", presentation)));

      WorkbookSheetResult.SheetLayout layout = workbook.sheet("Budget").layout();
      assertEquals(presentation.display(), layout.presentation().display());
      assertEquals(ExcelColorSnapshot.rgb("#225577"), layout.presentation().tabColor());
      assertEquals(presentation.outlineSummary(), layout.presentation().outlineSummary());
      assertEquals(presentation.sheetDefaults(), layout.presentation().sheetDefaults());
      assertEquals(
          List.of(
              new ExcelIgnoredError("A1:B2", List.of(ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT))),
          layout.presentation().ignoredErrors());
    }
  }

  @Test
  void appliesPivotTableCommandsAndSurfacesReadback() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      WorkbookCommandExecutor executor = new WorkbookCommandExecutor();

      executor.apply(
          workbook,
          new WorkbookSheetCommand.CreateSheet("Data"),
          new WorkbookSheetCommand.CreateSheet("Report"),
          new WorkbookCellCommand.SetRange(
              "Data",
              "A1:D5",
              List.of(
                  List.of(
                      ExcelCellValue.text("Region"),
                      ExcelCellValue.text("Stage"),
                      ExcelCellValue.text("Owner"),
                      ExcelCellValue.text("Amount")),
                  List.of(
                      ExcelCellValue.text("North"),
                      ExcelCellValue.text("Plan"),
                      ExcelCellValue.text("Ada"),
                      ExcelCellValue.number(10)),
                  List.of(
                      ExcelCellValue.text("North"),
                      ExcelCellValue.text("Do"),
                      ExcelCellValue.text("Ada"),
                      ExcelCellValue.number(15)),
                  List.of(
                      ExcelCellValue.text("South"),
                      ExcelCellValue.text("Plan"),
                      ExcelCellValue.text("Lin"),
                      ExcelCellValue.number(7)),
                  List.of(
                      ExcelCellValue.text("South"),
                      ExcelCellValue.text("Do"),
                      ExcelCellValue.text("Lin"),
                      ExcelCellValue.number(12)))),
          new WorkbookTabularCommand.SetPivotTable(
              new ExcelPivotTableDefinition(
                  "Ops Pivot",
                  "Report",
                  new ExcelPivotTableDefinition.Source.Range("Data", "A1:D5"),
                  new ExcelPivotTableDefinition.Anchor("A3"),
                  List.of("Region"),
                  List.of("Stage"),
                  List.of(),
                  List.of(
                      new ExcelPivotTableDefinition.DataField(
                          "Amount",
                          ExcelPivotDataConsolidateFunction.SUM,
                          "Total Amount",
                          "#,##0.00")))));

      WorkbookDrawingResult.PivotTablesResult pivotTables =
          assertInstanceOf(
              WorkbookDrawingResult.PivotTablesResult.class,
              new ExcelWorkbookIntrospector()
                  .execute(
                      workbook,
                      new WorkbookReadCommand.GetPivotTables(
                          "pivots", new ExcelPivotTableSelection.All())));
      ExcelPivotTableSnapshot.Supported pivot =
          assertInstanceOf(
              ExcelPivotTableSnapshot.Supported.class, pivotTables.pivotTables().getFirst());

      assertEquals("Ops Pivot", pivot.name());
      assertEquals("Amount", pivot.dataFields().getFirst().sourceColumnName());

      executor.apply(workbook, new WorkbookTabularCommand.DeletePivotTable("Ops Pivot", "Report"));

      WorkbookDrawingResult.PivotTablesResult afterDelete =
          assertInstanceOf(
              WorkbookDrawingResult.PivotTablesResult.class,
              new ExcelWorkbookIntrospector()
                  .execute(
                      workbook,
                      new WorkbookReadCommand.GetPivotTables(
                          "pivots", new ExcelPivotTableSelection.All())));
      assertEquals(List.of(), afterDelete.pivotTables());
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
          new WorkbookSheetCommand.CreateSheet("Ops"),
          new WorkbookDrawingCommand.SetPicture(
              "Ops",
              new ExcelPictureDefinition(
                  "OpsPicture",
                  new ExcelBinaryData(pngBytes),
                  ExcelPictureFormat.PNG,
                  firstAnchor,
                  "Queue preview")),
          new WorkbookDrawingCommand.SetShape(
              "Ops",
              new ExcelShapeDefinition(
                  "OpsShape",
                  ExcelAuthoredDrawingShapeKind.SIMPLE_SHAPE,
                  firstAnchor,
                  "rect",
                  "Queue")),
          new WorkbookDrawingCommand.SetEmbeddedObject(
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
          new WorkbookDrawingCommand.SetDrawingObjectAnchor("Ops", "OpsShape", movedAnchor),
          new WorkbookDrawingCommand.DeleteDrawingObject("Ops", "OpsPicture"));

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
          new WorkbookSheetCommand.SetWorkbookProtection(
              new ExcelWorkbookProtectionSettings(true, false, true, "secret", "review")));

      assertEquals(
          new ExcelWorkbookProtectionSnapshot(true, false, true, true, true),
          workbook.workbookProtection());

      executor.apply(workbook, new WorkbookSheetCommand.ClearWorkbookProtection());

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
          executor.apply(workbook, List.of(new WorkbookLayoutCommand.ClearPrintLayout("Budget"))));
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
          new WorkbookFormattingCommand.SetConditionalFormatting(
              "Ops",
              new ExcelConditionalFormattingBlockDefinition(
                  List.of("A1:A3"),
                  List.of(
                      new ExcelConditionalFormattingRule.FormulaRule(
                          "A1>0",
                          true,
                          new ExcelDifferentialStyle(
                              null, true, null, null, "#112233", null, null, null, null))))));

      WorkbookRuleResult.ConditionalFormattingResult initial =
          assertInstanceOf(
              WorkbookRuleResult.ConditionalFormattingResult.class,
              new ExcelWorkbookIntrospector()
                  .execute(
                      workbook,
                      new WorkbookReadCommand.GetConditionalFormatting(
                          "cf", "Ops", new ExcelRangeSelection.All())));
      assertEquals(1, initial.conditionalFormattingBlocks().size());

      executor.apply(
          workbook,
          new WorkbookFormattingCommand.ClearConditionalFormatting(
              "Ops", new ExcelRangeSelection.All()));

      WorkbookRuleResult.ConditionalFormattingResult cleared =
          assertInstanceOf(
              WorkbookRuleResult.ConditionalFormattingResult.class,
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
              new WorkbookSheetCommand.CreateSheet("Layout"),
              new WorkbookCellCommand.SetRange(
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
              new WorkbookStructureCommand.GroupRows("Layout", new ExcelRowSpan(1, 3), true),
              new WorkbookStructureCommand.SetRowVisibility("Layout", new ExcelRowSpan(5, 5), true),
              new WorkbookStructureCommand.GroupColumns("Layout", new ExcelColumnSpan(1, 3), true),
              new WorkbookStructureCommand.SetColumnVisibility(
                  "Layout", new ExcelColumnSpan(5, 5), true),
              new WorkbookSheetCommand.CreateSheet("Moves"),
              new WorkbookCellCommand.SetRange(
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
              new WorkbookStructureCommand.InsertRows("Moves", 1, 1),
              new WorkbookCellCommand.SetCell("Moves", "A2", ExcelCellValue.text("Spacer")),
              new WorkbookStructureCommand.ShiftRows("Moves", new ExcelRowSpan(2, 3), 1),
              new WorkbookStructureCommand.DeleteRows("Moves", new ExcelRowSpan(2, 2)),
              new WorkbookStructureCommand.InsertColumns("Moves", 1, 1),
              new WorkbookCellCommand.SetCell("Moves", "B1", ExcelCellValue.text("Pad")),
              new WorkbookStructureCommand.ShiftColumns("Moves", new ExcelColumnSpan(2, 4), 1),
              new WorkbookStructureCommand.DeleteColumns("Moves", new ExcelColumnSpan(2, 2))));

      assertEquals("Pad", workbook.sheet("Moves").text("B1"));
      assertEquals("Hosting", workbook.sheet("Moves").text("A3"));
      assertEquals(42.0, workbook.sheet("Moves").number("C3"));
      assertEquals("Beta", workbook.sheet("Moves").text("E4"));
      workbook.save(workbookPath);
    }

    WorkbookSheetResult.SheetLayout layout = XlsxRoundTrip.sheetLayout(workbookPath, "Layout");
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
              new WorkbookSheetCommand.CreateSheet("Alpha"),
              new WorkbookSheetCommand.CreateSheet("Beta"),
              new WorkbookCellCommand.SetCell("Alpha", "A1", ExcelCellValue.text("Live")),
              new WorkbookSheetCommand.CopySheet(
                  "Alpha", "Alpha Copy", new ExcelSheetCopyPosition.AtIndex(1)),
              new WorkbookSheetCommand.SetActiveSheet("Alpha"),
              new WorkbookSheetCommand.SetSelectedSheets(List.of("Alpha", "Beta")),
              new WorkbookSheetCommand.SetSheetVisibility("Beta", ExcelSheetVisibility.HIDDEN),
              new WorkbookSheetCommand.SetSheetProtection("Alpha", protectionSettings()),
              new WorkbookSheetCommand.ClearSheetProtection("Alpha")));

      assertEquals(List.of("Alpha", "Alpha Copy", "Beta"), workbook.sheetNames());
      assertEquals("Live", workbook.sheet("Alpha Copy").text("A1"));

      WorkbookCoreResult.WorkbookSummary.WithSheets summary =
          assertInstanceOf(
              WorkbookCoreResult.WorkbookSummary.WithSheets.class, workbook.workbookSummary());
      assertEquals("Alpha", summary.activeSheetName());
      assertEquals(List.of("Alpha"), summary.selectedSheetNames());
      assertEquals(ExcelSheetVisibility.HIDDEN, workbook.sheetSummary("Beta").visibility());
      assertInstanceOf(
          WorkbookSheetResult.SheetProtection.Unprotected.class,
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
          new WorkbookSheetCommand.CreateSheet("Layout"),
          new WorkbookCellCommand.SetRange(
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
          new WorkbookStructureCommand.GroupRows("Layout", new ExcelRowSpan(1, 3), true),
          new WorkbookStructureCommand.GroupColumns("Layout", new ExcelColumnSpan(1, 3), true),
          new WorkbookStructureCommand.UngroupRows("Layout", new ExcelRowSpan(1, 3)),
          new WorkbookStructureCommand.UngroupColumns("Layout", new ExcelColumnSpan(1, 3)));

      workbook.save(workbookPath);
    }

    WorkbookSheetResult.SheetLayout layout = XlsxRoundTrip.sheetLayout(workbookPath, "Layout");

    assertTrue(layout.rows().size() >= 4);
    assertFalse(layout.rows().get(1).hidden());
    assertEquals(0, layout.rows().get(1).outlineLevel());
    assertTrue(layout.columns().size() >= 4);
    assertFalse(layout.columns().get(1).hidden());
    assertEquals(0, layout.columns().get(1).outlineLevel());
  }

  @Test
  void rejectsUnexpectedCommandsInPrivateCommandFamilies() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      workbook.getOrCreateSheet("Budget");

      assertHelperRejects(
          () ->
              WorkbookCommandExecutor.applyWorkbookScopeCommand(
                  workbook,
                  new WorkbookAnnotationCommand.SetHyperlink(
                      "Budget", "A1", new ExcelHyperlink.Url("https://example.com"))));
      assertHelperRejects(
          () ->
              WorkbookCommandExecutor.applySheetStructureCommand(
                  workbook, new WorkbookSheetCommand.CreateSheet("Ops")));
      assertHelperRejects(
          () ->
              WorkbookCommandExecutor.applyCellValueCommand(
                  workbook,
                  new WorkbookSheetCommand.SetWorkbookProtection(
                      new ExcelWorkbookProtectionSettings(true, false, false, null, null))));
      assertHelperRejects(
          () ->
              WorkbookCommandExecutor.applyWorkbookMetadataCommand(
                  workbook, new WorkbookSheetCommand.CreateSheet("Ops")));
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

  private static void assertHelperRejects(Runnable invocation) {
    assertInstanceOf(
        IllegalStateException.class, assertThrows(IllegalStateException.class, invocation::run));
  }
}
