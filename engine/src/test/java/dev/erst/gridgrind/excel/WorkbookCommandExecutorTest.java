package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.io.IOException;
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
              new WorkbookCommand.FreezePanes("Budget", 1, 1, 1, 1),
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
          workbook.sheet("Budget").snapshotCell("A1").style().horizontalAlignment());
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
    assertEquals(
        new XlsxRoundTrip.FreezePaneState.Frozen(1, 1, 1, 1),
        XlsxRoundTrip.freezePaneState(workbookPath, "Budget"));
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
}
