package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.drawing.DrawingAnchorSpec;
import dev.erst.gridgrind.core.model.drawing.DrawingMarkerSpec;
import dev.erst.gridgrind.core.model.drawing.ResolvedComment;
import dev.erst.gridgrind.core.model.drawing.ResolvedEmbeddedObject;
import dev.erst.gridgrind.core.model.drawing.ResolvedPicture;
import dev.erst.gridgrind.core.model.drawing.ResolvedShape;
import dev.erst.gridgrind.core.model.formula.ResolvedFormula;
import dev.erst.gridgrind.core.model.print.HeaderFooterTextReport;
import dev.erst.gridgrind.core.model.print.PrintAreaSpec;
import dev.erst.gridgrind.core.model.print.PrintScalingSpec;
import dev.erst.gridgrind.core.model.print.PrintTitleColumnsSpec;
import dev.erst.gridgrind.core.model.print.PrintTitleRowsSpec;
import dev.erst.gridgrind.core.model.print.ResolvedPrintLayout;
import dev.erst.gridgrind.core.model.protection.ResolvedWorkbookProtection;
import dev.erst.gridgrind.core.model.protection.SheetProtectionSettings;
import dev.erst.gridgrind.core.model.protection.WorkbookProtectionReport;
import dev.erst.gridgrind.core.model.style.CellAlignmentPatch;
import dev.erst.gridgrind.core.model.style.CellColorReport;
import dev.erst.gridgrind.core.model.style.CellStylePatch;
import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.model.table.ResolvedTable;
import dev.erst.gridgrind.core.model.table.TableStyleSpec;
import dev.erst.gridgrind.core.model.validation.DataValidationRuleSpec;
import dev.erst.gridgrind.core.model.validation.ResolvedDataValidation;
import dev.erst.gridgrind.core.model.validation.ResolvedValidationErrorAlert;
import dev.erst.gridgrind.core.model.validation.ResolvedValidationPrompt;
import dev.erst.gridgrind.core.model.value.CellValue;
import dev.erst.gridgrind.core.source.BinaryContent;
import dev.erst.gridgrind.core.vocabulary.ExcelColumnSpan;
import dev.erst.gridgrind.core.vocabulary.ExcelDataValidationErrorStyle;
import dev.erst.gridgrind.core.vocabulary.ExcelDrawingAnchorBehavior;
import dev.erst.gridgrind.core.vocabulary.ExcelHorizontalAlignment;
import dev.erst.gridgrind.core.vocabulary.ExcelIgnoredErrorType;
import dev.erst.gridgrind.core.vocabulary.ExcelPictureFormat;
import dev.erst.gridgrind.core.vocabulary.ExcelPivotDataConsolidateFunction;
import dev.erst.gridgrind.core.vocabulary.ExcelPrintOrientation;
import dev.erst.gridgrind.core.vocabulary.ExcelRowSpan;
import dev.erst.gridgrind.core.vocabulary.ExcelSheetVisibility;
import dev.erst.gridgrind.core.vocabulary.ExcelVerticalAlignment;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import dev.erst.gridgrind.core.vocabulary.TableName;
import dev.erst.gridgrind.excel.drawing.ExcelDrawingObjectSnapshot;
import dev.erst.gridgrind.excel.pivot.ExcelPivotTableDefinition;
import dev.erst.gridgrind.excel.pivot.ExcelPivotTableSelection;
import dev.erst.gridgrind.excel.pivot.ExcelPivotTableSnapshot;
import dev.erst.gridgrind.excel.validation.ExcelDataValidationSnapshot;
import java.io.IOException;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Integration tests for WorkbookCommandExecutor applying commands to a workbook. */
class WorkbookCommandExecutorTest {
  @Test
  void appliesAllSupportedCommandTypesThroughVarargs() throws IOException {
    var workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-command-layout-");

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
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
                      List.of(new CellValue.Text("Item"), new CellValue.NumberValue(49.0)),
                      List.of(new CellValue.Text("Tax"), new CellValue.NumberValue(5.0)))),
              new WorkbookStructureCommand.MergeCells("Budget", "A1:B1"),
              new WorkbookFormattingCommand.ApplyStyle(
                  "Budget",
                  "A1:B1",
                  new CellStylePatch(
                      Optional.empty(),
                      Optional.of(
                          new CellAlignmentPatch(
                              Optional.empty(),
                              Optional.ofNullable(ExcelHorizontalAlignment.CENTER),
                              Optional.ofNullable(ExcelVerticalAlignment.CENTER),
                              Optional.empty(),
                              Optional.empty())),
                      Optional.empty(),
                      Optional.empty(),
                      Optional.empty(),
                      Optional.empty())),
              new WorkbookCellCommand.AppendRow(
                  "Budget",
                  List.of(new CellValue.Text("Total"), new ResolvedFormula("SUM(B1:B2)"))),
              new WorkbookCellCommand.ClearRange("Budget", "A2"),
              new WorkbookAnnotationCommand.SetHyperlink(
                  "Budget", "A1", new ExcelHyperlink.Document("Budget!B3")),
              new WorkbookAnnotationCommand.ClearHyperlink("Budget", "A1"),
              new WorkbookAnnotationCommand.SetComment(
                  "Budget", "B1", new ResolvedComment("Review", "GridGrind", false)),
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
                      List.of(new CellValue.Text("Owner"), new CellValue.Text("Task")),
                      List.of(new CellValue.Text("Ada"), new CellValue.Text("Queue")),
                      List.of(new CellValue.Text("Lin"), new CellValue.Text("Pack")))),
              new WorkbookCellCommand.SetRange(
                  "Ops",
                  "D1:E3",
                  List.of(
                      List.of(new CellValue.Text("Region"), new CellValue.Text("Desk")),
                      List.of(new CellValue.Text("North"), new CellValue.Text("A1")),
                      List.of(new CellValue.Text("South"), new CellValue.Text("B1")))),
              new WorkbookCellCommand.SetRange(
                  "Ops",
                  "G1:H3",
                  List.of(
                      List.of(new CellValue.Text("Stage"), new CellValue.Text("Team")),
                      List.of(new CellValue.Text("Review"), new CellValue.Text("Docs")),
                      List.of(new CellValue.Text("Ship"), new CellValue.Text("Ops")))),
              new WorkbookTabularCommand.SetAutofilter("Ops", "D1:E3"),
              new WorkbookTabularCommand.SetTable(
                  new ResolvedTable(
                      new TableName("Queue"),
                      new SheetName("Ops"),
                      new RangeRef("A1:B3"),
                      false,
                      new TableStyleSpec.None())),
              new WorkbookTabularCommand.SetTable(
                  new ResolvedTable(
                      new TableName("Stages"),
                      new SheetName("Ops"),
                      new RangeRef("G1:H3"),
                      false,
                      new TableStyleSpec.Named("TableStyleMedium2", false, false, true, false))),
              new WorkbookTabularCommand.DeleteTable("Stages", "Ops"),
              new WorkbookSheetCommand.RenameSheet("Archive", "History"),
              new WorkbookSheetCommand.MoveSheet("History", 0),
              new WorkbookSheetCommand.DeleteSheet("Scratch"),
              new WorkbookMetadataCommand.SetNamedRange(
                  new ExcelNamedRangeDefinition(
                      "BudgetTotal",
                      new ExcelNamedRangeScope.WorkbookScope(),
                      ExcelNamedRangeTarget.range("Budget", "B3"))),
              new WorkbookMetadataCommand.SetNamedRange(
                  new ExcelNamedRangeDefinition(
                      "HistoryCell",
                      new ExcelNamedRangeScope.SheetScope("History"),
                      ExcelNamedRangeTarget.range("History", "A1"))),
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
                  new ResolvedPrintLayout(
                      new PrintAreaSpec.Range(new RangeRef("A1:B3")),
                      ExcelPrintOrientation.LANDSCAPE,
                      new PrintScalingSpec.Fit(1, 0),
                      new PrintTitleRowsSpec.Band(0, 0),
                      new PrintTitleColumnsSpec.Band(0, 0),
                      new HeaderFooterTextReport("Budget", "", ""),
                      new HeaderFooterTextReport("", "Confidential", ""))),
              new WorkbookStructureCommand.UnmergeCells("Budget", "A1:B1")));
      workbook.formulas().markRecalculateOnOpen();
      assertEquals(List.of("History", "Budget", "Ops"), workbook.sheets().sheetNames());
      assertEquals("Item", workbook.sheet("Budget").cells().text("A1"));
      assertEquals(54.0, workbook.sheet("Budget").cells().number("B3"));
      assertTrue(
          workbook.sheet("Budget").cells().snapshotCell("A1").metadata().hyperlink().isEmpty());
      assertTrue(
          workbook.sheet("Budget").cells().snapshotCell("B1").metadata().comment().isEmpty());
      assertEquals(1, workbook.names().namedRangeCount());
      assertEquals(
          ExcelHorizontalAlignment.CENTER,
          workbook
              .sheet("Budget")
              .cells()
              .snapshotCell("A1")
              .style()
              .alignment()
              .horizontalAlignment());
      assertEquals("BLANK", workbook.sheet("Budget").cells().snapshotCell("A2").type());
      assertThrows(SheetNotFoundException.class, () -> workbook.sheet("Scratch"));
      assertTrue(workbook.formulas().recalculateOnOpenEnabled());
      WorkbookRuleResult.AutofiltersResult autofilters =
          assertInstanceOf(
              WorkbookRuleResult.AutofiltersResult.class,
              new ExcelWorkbookIntrospector()
                  .execute(
                      workbook,
                      new WorkbookReadCommand.GetAutofilters(new StepId("autofilters"), "Ops")));
      WorkbookRuleResult.TablesResult tables =
          assertInstanceOf(
              WorkbookRuleResult.TablesResult.class,
              new ExcelWorkbookIntrospector()
                  .execute(
                      workbook,
                      new WorkbookReadCommand.GetTables(
                          new StepId("tables"), new ExcelTableSelection.All())));
      assertEquals(
          List.of(
              new ExcelAutofilterSnapshot.SheetOwned("D1:E3"),
              new ExcelAutofilterSnapshot.TableOwned("A1:B3", "Queue")),
          autofilters.autofilters());
      assertEquals(
          List.of("Queue"), tables.tables().stream().map(ExcelTableSnapshot::name).toList());
      workbook
          .persistence()
          .save(
              workbookPath,
              dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
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
                ExcelNamedRangeTarget.range("History", "A1"))),
        XlsxRoundTrip.namedRanges(workbookPath));
    try (ExcelWorkbook reopened =
        ExcelWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      WorkbookRuleResult.AutofiltersResult autofilters =
          assertInstanceOf(
              WorkbookRuleResult.AutofiltersResult.class,
              new ExcelWorkbookIntrospector()
                  .execute(
                      reopened,
                      new WorkbookReadCommand.GetAutofilters(new StepId("autofilters"), "Ops")));
      WorkbookRuleResult.TablesResult tables =
          assertInstanceOf(
              WorkbookRuleResult.TablesResult.class,
              new ExcelWorkbookIntrospector()
                  .execute(
                      reopened,
                      new WorkbookReadCommand.GetTables(
                          new StepId("tables"), new ExcelTableSelection.All())));
      assertEquals(
          List.of(
              new ExcelAutofilterSnapshot.SheetOwned("D1:E3"),
              new ExcelAutofilterSnapshot.TableOwned("A1:B3", "Queue")),
          autofilters.autofilters());
      assertEquals(1, tables.tables().size());
      assertEquals("Queue", tables.tables().getFirst().name());
      assertTrue(tables.tables().getFirst().behavior().hasAutofilter());
    }
  }

  @Test
  void rejectsMutationCommandsWhenSheetDoesNotExist() throws IOException {
    WorkbookCommandExecutor executor = new WorkbookCommandExecutor();

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      assertThrows(
          SheetNotFoundException.class,
          () ->
              executor.apply(
                  workbook,
                  new WorkbookCellCommand.SetCell("Missing", "A1", new CellValue.Text("x"))));
      assertThrows(
          SheetNotFoundException.class,
          () ->
              executor.apply(
                  workbook,
                  new WorkbookCellCommand.SetRange(
                      "Missing", "A1", List.of(List.of(new CellValue.Text("x"))))));
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
                      "Missing", "A1", new ResolvedComment("Review", "GridGrind", false))));
      assertThrows(
          SheetNotFoundException.class,
          () ->
              executor.apply(
                  workbook,
                  new WorkbookFormattingCommand.ApplyStyle(
                      "Missing",
                      "A1",
                      new CellStylePatch(
                          Optional.empty(),
                          Optional.of(
                              new CellAlignmentPatch(
                                  Optional.empty(),
                                  Optional.ofNullable(ExcelHorizontalAlignment.CENTER),
                                  Optional.ofNullable(ExcelVerticalAlignment.TOP),
                                  Optional.empty(),
                                  Optional.empty())),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty(),
                          Optional.empty()))));
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
                                  Optional.of(
                                      new ExcelDifferentialStyle(
                                          Optional.empty(),
                                          Optional.of(true),
                                          Optional.empty(),
                                          Optional.empty(),
                                          Optional.empty(),
                                          Optional.empty(),
                                          Optional.empty(),
                                          Optional.empty(),
                                          Optional.empty()))))))));
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
                      new ResolvedTable(
                          new TableName("Queue"),
                          new SheetName("Missing"),
                          new RangeRef("A1:B2"),
                          false,
                          new TableStyleSpec.None()))));
      assertThrows(
          IllegalArgumentException.class,
          () ->
              executor.apply(workbook, new WorkbookTabularCommand.DeleteTable("Queue", "Missing")));
      assertThrows(
          SheetNotFoundException.class,
          () ->
              executor.apply(
                  workbook,
                  new WorkbookCellCommand.AppendRow("Missing", List.of(new CellValue.Text("x")))));
    }
  }

  @Test
  void appliesSheetPresentationThroughSheetStructureDispatch() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      WorkbookCommandExecutor executor = new WorkbookCommandExecutor();
      workbook.getOrCreateSheet("Budget");
      ExcelSheetPresentation presentation =
          new ExcelSheetPresentation(
              new ExcelSheetDisplay(false, false, false, true, true),
              Optional.of(ColorSpec.rgb(new RgbHex("#225577"))),
              new ExcelSheetOutlineSummary(false, false),
              new ExcelSheetDefaults(12, 19.5d),
              List.of(
                  new ExcelIgnoredError(
                      "B2:A1", List.of(ExcelIgnoredErrorType.NUMBER_STORED_AS_TEXT))));

      assertSame(
          workbook,
          executor.apply(
              workbook, new WorkbookLayoutCommand.SetSheetPresentation("Budget", presentation)));

      WorkbookSheetResult.SheetLayout layout = workbook.sheet("Budget").layout().snapshot();
      assertEquals(presentation.display(), layout.presentation().display());
      assertEquals(
          Optional.of(CellColorReport.rgb(new RgbHex("#225577"))),
          layout.presentation().tabColor());
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
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
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
                      new CellValue.Text("Region"),
                      new CellValue.Text("Stage"),
                      new CellValue.Text("Owner"),
                      new CellValue.Text("Amount")),
                  List.of(
                      new CellValue.Text("North"),
                      new CellValue.Text("Plan"),
                      new CellValue.Text("Ada"),
                      new CellValue.NumberValue(10)),
                  List.of(
                      new CellValue.Text("North"),
                      new CellValue.Text("Do"),
                      new CellValue.Text("Ada"),
                      new CellValue.NumberValue(15)),
                  List.of(
                      new CellValue.Text("South"),
                      new CellValue.Text("Plan"),
                      new CellValue.Text("Lin"),
                      new CellValue.NumberValue(7)),
                  List.of(
                      new CellValue.Text("South"),
                      new CellValue.Text("Do"),
                      new CellValue.Text("Lin"),
                      new CellValue.NumberValue(12)))),
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
                          Optional.of("#,##0.00"))))));

      WorkbookDrawingResult.PivotTablesResult pivotTables =
          assertInstanceOf(
              WorkbookDrawingResult.PivotTablesResult.class,
              new ExcelWorkbookIntrospector()
                  .execute(
                      workbook,
                      new WorkbookReadCommand.GetPivotTables(
                          new StepId("pivots"), new ExcelPivotTableSelection.All())));
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
                          new StepId("pivots"), new ExcelPivotTableSelection.All())));
      assertEquals(List.of(), afterDelete.pivotTables());
    }
  }

  @Test
  void appliesDrawingCommandsThroughTheWorkbookCommandSurface() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      WorkbookCommandExecutor executor = new WorkbookCommandExecutor();
      DrawingAnchorSpec.TwoCell firstAnchor =
          new DrawingAnchorSpec.TwoCell(
              new DrawingMarkerSpec(1, 1, 0, 0),
              new DrawingMarkerSpec(4, 6, 0, 0),
              ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE);
      DrawingAnchorSpec.TwoCell movedAnchor =
          new DrawingAnchorSpec.TwoCell(
              new DrawingMarkerSpec(6, 2, 0, 0),
              new DrawingMarkerSpec(9, 7, 0, 0),
              ExcelDrawingAnchorBehavior.MOVE_DONT_RESIZE);
      byte[] pngBytes =
          java.util.Base64.getDecoder()
              .decode(
                  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVQI12P4//8/AAX+Av7czFnnAAAAAElFTkSuQmCC");

      executor.apply(
          workbook,
          new WorkbookSheetCommand.CreateSheet("Ops"),
          new WorkbookDrawingCommand.SetPicture(
              "Ops",
              new ResolvedPicture(
                  "OpsPicture",
                  new BinaryContent(pngBytes),
                  ExcelPictureFormat.PNG,
                  firstAnchor,
                  Optional.of("Queue preview"))),
          new WorkbookDrawingCommand.SetShape(
              "Ops",
              new ResolvedShape.SimpleShape("OpsShape", firstAnchor, "rect", Optional.of("Queue"))),
          new WorkbookDrawingCommand.SetEmbeddedObject(
              "Ops",
              new ResolvedEmbeddedObject(
                  "OpsEmbed",
                  "Payload",
                  "payload.txt",
                  "payload.txt",
                  new BinaryContent("payload".getBytes(java.nio.charset.StandardCharsets.UTF_8)),
                  ExcelPictureFormat.PNG,
                  new BinaryContent(pngBytes),
                  firstAnchor)),
          new WorkbookDrawingCommand.SetDrawingObjectAnchor("Ops", "OpsShape", movedAnchor),
          new WorkbookDrawingCommand.DeleteDrawingObject("Ops", "OpsPicture"));

      List<ExcelDrawingObjectSnapshot> drawingObjects =
          workbook.sheet("Ops").drawings().drawingObjects();
      assertEquals(
          List.of("OpsShape", "OpsEmbed"),
          drawingObjects.stream().map(ExcelDrawingObjectSnapshot::name).toList());
      assertEquals(
          DrawingGeometryFixtures.report(movedAnchor),
          assertInstanceOf(ExcelDrawingObjectSnapshot.Shape.class, drawingObjects.getFirst())
              .anchor());
    }
  }

  @Test
  void appliesWorkbookProtectionCommands() throws IOException {
    WorkbookCommandExecutor executor = new WorkbookCommandExecutor();

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Budget");

      executor.apply(
          workbook,
          new WorkbookSheetCommand.SetWorkbookProtection(
              new ResolvedWorkbookProtection(
                  true, false, true, Optional.of("secret"), Optional.of("review"))));

      assertEquals(
          new WorkbookProtectionReport(true, false, true, true, true),
          workbook.workbookProtection());

      executor.apply(workbook, new WorkbookSheetCommand.ClearWorkbookProtection());

      assertEquals(
          new WorkbookProtectionReport(false, false, false, false, false),
          workbook.workbookProtection());
    }
  }

  @Test
  void appliesClearPrintLayoutThroughIterableCommands() throws IOException {
    WorkbookCommandExecutor executor = new WorkbookCommandExecutor();

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      workbook.getOrCreateSheet("Budget");
      workbook
          .sheet("Budget")
          .layout()
          .setPrintLayout(
              new ResolvedPrintLayout(
                  new PrintAreaSpec.Range(new RangeRef("A1:B3")),
                  ExcelPrintOrientation.LANDSCAPE,
                  new PrintScalingSpec.Fit(1, 0),
                  new PrintTitleRowsSpec.Band(0, 0),
                  new PrintTitleColumnsSpec.Band(0, 0),
                  new HeaderFooterTextReport("Budget", "", ""),
                  new HeaderFooterTextReport("", "Confidential", "")));

      assertSame(
          workbook,
          executor.apply(workbook, List.of(new WorkbookLayoutCommand.ClearPrintLayout("Budget"))));
      ResolvedPrintLayout clearedPrintLayout = workbook.sheet("Budget").layout().printLayout();
      assertEquals(new PrintAreaSpec.None(), clearedPrintLayout.printArea());
      assertEquals(ExcelPrintOrientation.PORTRAIT, clearedPrintLayout.orientation());
      assertEquals(new PrintScalingSpec.Automatic(), clearedPrintLayout.scaling());
      assertEquals(new PrintTitleRowsSpec.None(), clearedPrintLayout.repeatingRows());
      assertEquals(new PrintTitleColumnsSpec.None(), clearedPrintLayout.repeatingColumns());
      assertEquals(HeaderFooterTextReport.blank(), clearedPrintLayout.header());
      assertEquals(HeaderFooterTextReport.blank(), clearedPrintLayout.footer());
    }
  }

  @Test
  void appliesConditionalFormattingCommandsThroughExecutor() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
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
                          Optional.of(
                              new ExcelDifferentialStyle(
                                  Optional.empty(),
                                  Optional.of(true),
                                  Optional.empty(),
                                  Optional.empty(),
                                  Optional.of(ColorSpec.rgb(new RgbHex("#112233"))),
                                  Optional.empty(),
                                  Optional.empty(),
                                  Optional.empty(),
                                  Optional.empty())))))));

      WorkbookRuleResult.ConditionalFormattingResult initial =
          assertInstanceOf(
              WorkbookRuleResult.ConditionalFormattingResult.class,
              new ExcelWorkbookIntrospector()
                  .execute(
                      workbook,
                      new WorkbookReadCommand.GetConditionalFormatting(
                          new StepId("cf"), "Ops", new ExcelRangeSelection.All())));
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
                          new StepId("cf"), "Ops", new ExcelRangeSelection.All())));
      assertEquals(List.of(), cleared.conditionalFormattingBlocks());
    }
  }

  @Test
  void appliesRowAndColumnStructuralCommandsThroughExecutor() throws IOException {
    Path workbookPath = XlsxRoundTrip.newWorkbookPath("gridgrind-structural-commands-");

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
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
                          new CellValue.Text("Item"),
                          new CellValue.Text("Qty"),
                          new CellValue.Text("Status"),
                          new CellValue.Text("Note"),
                          new CellValue.Text("Owner"),
                          new CellValue.Text("Flag")),
                      List.of(
                          new CellValue.Text("Hosting"),
                          new CellValue.NumberValue(42.0),
                          new CellValue.Text("Open"),
                          new CellValue.Text("Alpha"),
                          new CellValue.Text("Ada"),
                          new CellValue.Text("Y")),
                      List.of(
                          new CellValue.Text("Support"),
                          new CellValue.NumberValue(84.0),
                          new CellValue.Text("Closed"),
                          new CellValue.Text("Beta"),
                          new CellValue.Text("Lin"),
                          new CellValue.Text("N")),
                      List.of(
                          new CellValue.Text("Ops"),
                          new CellValue.NumberValue(168.0),
                          new CellValue.Text("Open"),
                          new CellValue.Text("Gamma"),
                          new CellValue.Text("Bea"),
                          new CellValue.Text("Y")),
                      List.of(
                          new CellValue.Text("QA"),
                          new CellValue.NumberValue(21.0),
                          new CellValue.Text("Queued"),
                          new CellValue.Text("Delta"),
                          new CellValue.Text("Kai"),
                          new CellValue.Text("N")),
                      List.of(
                          new CellValue.Text("Infra"),
                          new CellValue.NumberValue(7.0),
                          new CellValue.Text("Done"),
                          new CellValue.Text("Epsilon"),
                          new CellValue.Text("Mia"),
                          new CellValue.Text("Y")))),
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
                          new CellValue.Text("Item"),
                          new CellValue.Text("Qty"),
                          new CellValue.Text("Status"),
                          new CellValue.Text("Note")),
                      List.of(
                          new CellValue.Text("Hosting"),
                          new CellValue.NumberValue(42.0),
                          new CellValue.Text("Open"),
                          new CellValue.Text("Alpha")),
                      List.of(
                          new CellValue.Text("Support"),
                          new CellValue.NumberValue(84.0),
                          new CellValue.Text("Closed"),
                          new CellValue.Text("Beta")))),
              new WorkbookStructureCommand.InsertRows("Moves", 1, 1),
              new WorkbookCellCommand.SetCell("Moves", "A2", new CellValue.Text("Spacer")),
              new WorkbookStructureCommand.ShiftRows("Moves", new ExcelRowSpan(2, 3), 1),
              new WorkbookStructureCommand.DeleteRows("Moves", new ExcelRowSpan(2, 2)),
              new WorkbookStructureCommand.InsertColumns("Moves", 1, 1),
              new WorkbookCellCommand.SetCell("Moves", "B1", new CellValue.Text("Pad")),
              new WorkbookStructureCommand.ShiftColumns("Moves", new ExcelColumnSpan(2, 4), 1),
              new WorkbookStructureCommand.DeleteColumns("Moves", new ExcelColumnSpan(2, 2))));

      assertEquals("Pad", workbook.sheet("Moves").cells().text("B1"));
      assertEquals("Hosting", workbook.sheet("Moves").cells().text("A3"));
      assertEquals(42.0, workbook.sheet("Moves").cells().number("C3"));
      assertEquals("Beta", workbook.sheet("Moves").cells().text("E4"));
      workbook
          .persistence()
          .save(
              workbookPath,
              dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
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

    try (ExcelWorkbook reopened =
        ExcelWorkbooks.open(workbookPath, ExcelTempFileFactoryTestSupport.tempFileFactory())) {
      assertEquals("Pad", reopened.sheet("Moves").cells().text("B1"));
      assertEquals("Hosting", reopened.sheet("Moves").cells().text("A3"));
      assertEquals(42.0, reopened.sheet("Moves").cells().number("C3"));
      assertEquals("Beta", reopened.sheet("Moves").cells().text("E4"));
    }
  }

  @Test
  void appliesSheetManagementCommandsThroughExecutor() throws IOException {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      WorkbookCommandExecutor executor = new WorkbookCommandExecutor();

      assertSame(
          workbook,
          executor.apply(
              workbook,
              new WorkbookSheetCommand.CreateSheet("Alpha"),
              new WorkbookSheetCommand.CreateSheet("Beta"),
              new WorkbookCellCommand.SetCell("Alpha", "A1", new CellValue.Text("Live")),
              new WorkbookSheetCommand.CopySheet(
                  "Alpha", "Alpha Copy", new ExcelSheetCopyPosition.AtIndex(1)),
              new WorkbookSheetCommand.SetActiveSheet("Alpha"),
              new WorkbookSheetCommand.SetSelectedSheets(List.of("Alpha", "Beta")),
              new WorkbookSheetCommand.SetSheetVisibility("Beta", ExcelSheetVisibility.HIDDEN),
              new WorkbookSheetCommand.SetSheetProtection("Alpha", protectionSettings()),
              new WorkbookSheetCommand.ClearSheetProtection("Alpha")));

      assertEquals(List.of("Alpha", "Alpha Copy", "Beta"), workbook.sheets().sheetNames());
      assertEquals("Live", workbook.sheet("Alpha Copy").cells().text("A1"));

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

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      WorkbookCommandExecutor executor = new WorkbookCommandExecutor();

      executor.apply(
          workbook,
          new WorkbookSheetCommand.CreateSheet("Layout"),
          new WorkbookCellCommand.SetRange(
              "Layout",
              "A1:D4",
              List.of(
                  List.of(
                      new CellValue.Text("Item"),
                      new CellValue.Text("Qty"),
                      new CellValue.Text("Status"),
                      new CellValue.Text("Owner")),
                  List.of(
                      new CellValue.Text("Hosting"),
                      new CellValue.NumberValue(42.0),
                      new CellValue.Text("Open"),
                      new CellValue.Text("Ada")),
                  List.of(
                      new CellValue.Text("Support"),
                      new CellValue.NumberValue(84.0),
                      new CellValue.Text("Closed"),
                      new CellValue.Text("Lin")),
                  List.of(
                      new CellValue.Text("Ops"),
                      new CellValue.NumberValue(7.0),
                      new CellValue.Text("Done"),
                      new CellValue.Text("Mia")))),
          new WorkbookStructureCommand.GroupRows("Layout", new ExcelRowSpan(1, 3), true),
          new WorkbookStructureCommand.GroupColumns("Layout", new ExcelColumnSpan(1, 3), true),
          new WorkbookStructureCommand.UngroupRows("Layout", new ExcelRowSpan(1, 3)),
          new WorkbookStructureCommand.UngroupColumns("Layout", new ExcelColumnSpan(1, 3)));

      workbook
          .persistence()
          .save(
              workbookPath,
              dev.erst.gridgrind.excel.WorkbookArtifactWriteDisposition.REPLACE_EXISTING,
              ExcelTempFileFactoryTestSupport.tempFileFactory());
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
  void routesDedicatedHelperFamiliesAcrossWorkbookSeams() throws Exception {
    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
      WorkbookCommandExecutor.applyWorkbookScopeCommand(
          workbook, new WorkbookSheetCommand.CreateSheet("Budget"));
      WorkbookCommandExecutor.applyWorkbookScopeCommand(
          workbook, new WorkbookSheetCommand.CreateSheet("Ops"));
      assertEquals(List.of("Budget", "Ops"), workbook.sheets().sheetNames());

      WorkbookCommandExecutor.applyCellValueCommand(
          workbook,
          new WorkbookCellCommand.SetRange(
              "Budget",
              "A1:B2",
              List.of(
                  List.of(
                      new CellValue.Text("Extremely long workflow heading"),
                      new CellValue.Text("State")),
                  List.of(new CellValue.Text("Ada"), new CellValue.Text("Queued")))));
      WorkbookCommandExecutor.applySheetStructureCommand(
          workbook, new WorkbookStructureCommand.SetColumnWidth("Budget", 0, 0, 4.0));
      WorkbookCommandExecutor.applySheetLayoutCommand(
          workbook, new WorkbookLayoutCommand.SetSheetZoom("Budget", 135));

      double widthBeforeAutoSize =
          workbook.sheet("Budget").layout().snapshot().columns().getFirst().widthCharacters();
      WorkbookCommandExecutor.applySheetLayoutCommand(
          workbook, new WorkbookLayoutCommand.AutoSizeColumns("Budget"));

      assertEquals(135, workbook.sheet("Budget").layout().snapshot().zoomPercent());
      assertTrue(
          workbook.sheet("Budget").layout().snapshot().columns().getFirst().widthCharacters()
              > widthBeforeAutoSize);

      WorkbookCommandExecutor.applyAnnotationCommand(
          workbook,
          new WorkbookAnnotationCommand.SetComment(
              "Budget", "A1", new ResolvedComment("Review", "GridGrind", false)));
      WorkbookCommandExecutor.applyMetadataCommand(
          workbook,
          new WorkbookMetadataCommand.SetNamedRange(
              new ExcelNamedRangeDefinition(
                  "BudgetStatus",
                  new ExcelNamedRangeScope.WorkbookScope(),
                  ExcelNamedRangeTarget.range("Budget", "B2"))));
      WorkbookCommandExecutor.applyDrawingCommand(
          workbook,
          new WorkbookDrawingCommand.SetShape(
              "Budget",
              new ResolvedShape.SimpleShape(
                  "BudgetShape",
                  new DrawingAnchorSpec.TwoCell(
                      new DrawingMarkerSpec(0, 0, 0, 0),
                      new DrawingMarkerSpec(2, 3, 0, 0),
                      ExcelDrawingAnchorBehavior.MOVE_AND_RESIZE),
                  "rect",
                  Optional.of("Queue"))));
      WorkbookCommandExecutor.applyFormattingCommand(
          workbook,
          new WorkbookFormattingCommand.ApplyStyle(
              "Budget",
              "A1",
              new CellStylePatch(
                  Optional.empty(),
                  Optional.of(
                      new CellAlignmentPatch(
                          Optional.empty(),
                          Optional.ofNullable(ExcelHorizontalAlignment.CENTER),
                          Optional.ofNullable(ExcelVerticalAlignment.CENTER),
                          Optional.empty(),
                          Optional.empty())),
                  Optional.empty(),
                  Optional.empty(),
                  Optional.empty(),
                  Optional.empty())));
      WorkbookCommandExecutor.applyTabularCommand(
          workbook, new WorkbookTabularCommand.SetAutofilter("Budget", "A1:B2"));

      assertEquals(
          new ResolvedComment("Review", "GridGrind", false),
          workbook
              .sheet("Budget")
              .cells()
              .snapshotCell("A1")
              .metadata()
              .comment()
              .orElseThrow()
              .toPlainComment());
      assertEquals(1, workbook.names().namedRangeCount());
      assertEquals(1, workbook.sheet("Budget").drawings().drawingObjects().size());
      assertEquals(
          ExcelHorizontalAlignment.CENTER,
          workbook
              .sheet("Budget")
              .cells()
              .snapshotCell("A1")
              .style()
              .alignment()
              .horizontalAlignment());
      WorkbookRuleResult.AutofiltersResult autofilters =
          assertInstanceOf(
              WorkbookRuleResult.AutofiltersResult.class,
              new ExcelWorkbookIntrospector()
                  .execute(
                      workbook,
                      new WorkbookReadCommand.GetAutofilters(new StepId("autofilters"), "Budget")));
      assertEquals(
          List.of(new ExcelAutofilterSnapshot.SheetOwned("A1:B2")), autofilters.autofilters());
    }
  }

  @Test
  void validatesNullWorkbooksCommandsAndCommandEntries() throws IOException {
    WorkbookCommandExecutor executor = new WorkbookCommandExecutor();

    try (ExcelWorkbook workbook = ExcelWorkbooks.create()) {
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

  private static ResolvedDataValidation validationDefinition() {
    return new ResolvedDataValidation(
        new DataValidationRuleSpec.ExplicitList(List.of("Queued", "Done")),
        true,
        false,
        Optional.of(new ResolvedValidationPrompt("Status", "Pick one workflow state.", true)),
        Optional.of(
            new ResolvedValidationErrorAlert(
                ExcelDataValidationErrorStyle.STOP,
                "Invalid status",
                "Use one of the allowed values.",
                true)));
  }

  private static SheetProtectionSettings protectionSettings() {
    return new SheetProtectionSettings(
        true, false, true, false, true, false, true, false, true, false, true, false, true, false,
        true);
  }
}
