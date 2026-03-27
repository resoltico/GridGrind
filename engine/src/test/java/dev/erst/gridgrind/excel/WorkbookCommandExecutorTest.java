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
      assertEquals(List.of("History", "Budget"), workbook.sheetNames());
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
            new ExcelNamedRangeSnapshot.RangeSnapshot(
                "HistoryCell",
                new ExcelNamedRangeScope.SheetScope("History"),
                "History!$A$1",
                new ExcelNamedRangeTarget("History", "A1"))),
        XlsxRoundTrip.namedRanges(workbookPath));
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
}
