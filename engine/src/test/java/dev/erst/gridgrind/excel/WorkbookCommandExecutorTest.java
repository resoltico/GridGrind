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
    try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
      WorkbookCommandExecutor executor = new WorkbookCommandExecutor();

      assertSame(
          workbook,
          executor.apply(
              workbook,
              new WorkbookCommand.CreateSheet("Budget"),
              new WorkbookCommand.SetRange(
                  "Budget",
                  "A1:B2",
                  List.of(
                      List.of(ExcelCellValue.text("Item"), ExcelCellValue.number(49.0)),
                      List.of(ExcelCellValue.text("Tax"), ExcelCellValue.number(5.0)))),
              new WorkbookCommand.ApplyStyle(
                  "Budget",
                  "A1:B1",
                  ExcelCellStyle.alignment(
                      ExcelHorizontalAlignment.CENTER, ExcelVerticalAlignment.CENTER)),
              new WorkbookCommand.AppendRow(
                  "Budget",
                  List.of(ExcelCellValue.text("Total"), ExcelCellValue.formula("SUM(B1:B2)"))),
              new WorkbookCommand.ClearRange("Budget", "A2"),
              new WorkbookCommand.AutoSizeColumns("Budget"),
              new WorkbookCommand.EvaluateAllFormulas(),
              new WorkbookCommand.ForceFormulaRecalculationOnOpen()));
      assertEquals("Item", workbook.sheet("Budget").text("A1"));
      assertEquals(54.0, workbook.sheet("Budget").number("B3"));
      assertEquals(
          ExcelHorizontalAlignment.CENTER,
          workbook.sheet("Budget").snapshotCell("A1").style().horizontalAlignment());
      assertEquals("BLANK", workbook.sheet("Budget").snapshotCell("A2").effectiveType());
      assertTrue(workbook.forceFormulaRecalculationOnOpenEnabled());
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
}
