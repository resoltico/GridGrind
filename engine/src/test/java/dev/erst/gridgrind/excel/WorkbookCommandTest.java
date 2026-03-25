package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for WorkbookCommand sealed interface record construction. */
class WorkbookCommandTest {
  @Test
  void createsSupportedCommandsAndCopiesCollections() {
    List<ExcelCellValue> values = new ArrayList<>(List.of(ExcelCellValue.text("Item")));
    List<List<ExcelCellValue>> rows =
        new ArrayList<>(
            List.of(
                new ArrayList<>(List.of(ExcelCellValue.text("Item"), ExcelCellValue.number(49.0))),
                new ArrayList<>(List.of(ExcelCellValue.text("Tax"), ExcelCellValue.number(10.0)))));
    ExcelCellStyle style =
        new ExcelCellStyle(
            "#,##0.00",
            true,
            false,
            true,
            ExcelHorizontalAlignment.RIGHT,
            ExcelVerticalAlignment.CENTER);

    WorkbookCommand.CreateSheet createSheet = new WorkbookCommand.CreateSheet("Budget");
    WorkbookCommand.RenameSheet renameSheet = new WorkbookCommand.RenameSheet("Budget", "Summary");
    WorkbookCommand.DeleteSheet deleteSheet = new WorkbookCommand.DeleteSheet("Archive");
    WorkbookCommand.MoveSheet moveSheet = new WorkbookCommand.MoveSheet("Budget", 1);
    WorkbookCommand.MergeCells mergeCells = new WorkbookCommand.MergeCells("Budget", "A1:B2");
    WorkbookCommand.UnmergeCells unmergeCells = new WorkbookCommand.UnmergeCells("Budget", "A1:B2");
    WorkbookCommand.SetColumnWidth setColumnWidth =
        new WorkbookCommand.SetColumnWidth("Budget", 0, 1, 16.0);
    WorkbookCommand.SetRowHeight setRowHeight =
        new WorkbookCommand.SetRowHeight("Budget", 0, 2, 28.5);
    WorkbookCommand.FreezePanes freezePanes = new WorkbookCommand.FreezePanes("Budget", 1, 2, 1, 2);
    WorkbookCommand.SetCell setCell =
        new WorkbookCommand.SetCell("Budget", "A1", ExcelCellValue.date(LocalDate.of(2026, 3, 23)));
    WorkbookCommand.SetRange setRange = new WorkbookCommand.SetRange("Budget", "A1:B2", rows);
    WorkbookCommand.ClearRange clearRange = new WorkbookCommand.ClearRange("Budget", "C1:C2");
    WorkbookCommand.ApplyStyle applyStyle =
        new WorkbookCommand.ApplyStyle("Budget", "A1:B1", style);
    WorkbookCommand.AppendRow appendRow = new WorkbookCommand.AppendRow("Budget", values);
    WorkbookCommand.AutoSizeColumns autoSizeColumns = new WorkbookCommand.AutoSizeColumns("Budget");
    WorkbookCommand.EvaluateAllFormulas evaluate = new WorkbookCommand.EvaluateAllFormulas();
    WorkbookCommand.ForceFormulaRecalculationOnOpen recalc =
        new WorkbookCommand.ForceFormulaRecalculationOnOpen();

    values.clear();
    rows.clear();

    assertEquals("Budget", createSheet.sheetName());
    assertEquals("Summary", renameSheet.newSheetName());
    assertEquals("Archive", deleteSheet.sheetName());
    assertEquals(1, moveSheet.targetIndex());
    assertEquals("A1:B2", mergeCells.range());
    assertEquals("A1:B2", unmergeCells.range());
    assertEquals(16.0, setColumnWidth.widthCharacters());
    assertEquals(28.5, setRowHeight.heightPoints());
    assertEquals(2, freezePanes.topRow());
    assertEquals("A1", setCell.address());
    assertEquals("A1:B2", setRange.range());
    assertEquals(2, setRange.rows().size());
    assertEquals("C1:C2", clearRange.range());
    assertEquals(style, applyStyle.style());
    assertEquals(1, appendRow.values().size());
    assertEquals("Budget", autoSizeColumns.sheetName());
    assertNotNull(evaluate);
    assertNotNull(recalc);
  }

  @Test
  void validatesSheetAndLayoutCommandInputs() {
    assertThrows(NullPointerException.class, () -> new WorkbookCommand.CreateSheet(null));
    assertThrows(IllegalArgumentException.class, () -> new WorkbookCommand.CreateSheet(" "));
    assertThrows(NullPointerException.class, () -> new WorkbookCommand.RenameSheet(null, "New"));
    assertThrows(IllegalArgumentException.class, () -> new WorkbookCommand.RenameSheet(" ", "New"));
    assertThrows(NullPointerException.class, () -> new WorkbookCommand.RenameSheet("Budget", null));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCommand.RenameSheet("Budget", " "));
    assertThrows(NullPointerException.class, () -> new WorkbookCommand.DeleteSheet(null));
    assertThrows(IllegalArgumentException.class, () -> new WorkbookCommand.DeleteSheet(" "));
    assertThrows(NullPointerException.class, () -> new WorkbookCommand.MoveSheet(null, 0));
    assertThrows(IllegalArgumentException.class, () -> new WorkbookCommand.MoveSheet(" ", 0));
    assertThrows(IllegalArgumentException.class, () -> new WorkbookCommand.MoveSheet("Budget", -1));
    assertThrows(NullPointerException.class, () -> new WorkbookCommand.MergeCells(null, "A1:B2"));
    assertThrows(IllegalArgumentException.class, () -> new WorkbookCommand.MergeCells(" ", "A1"));
    assertThrows(NullPointerException.class, () -> new WorkbookCommand.MergeCells("Budget", null));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCommand.MergeCells("Budget", " "));
    assertThrows(
        NullPointerException.class, () -> new WorkbookCommand.UnmergeCells("Budget", null));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCommand.UnmergeCells(" ", "A1:B2"));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCommand.UnmergeCells("Budget", " "));
    assertThrows(
        NullPointerException.class, () -> new WorkbookCommand.SetColumnWidth(null, 0, 0, 16.0));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCommand.SetColumnWidth(" ", 0, 0, 16.0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetColumnWidth("Budget", -1, 0, 16.0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetColumnWidth("Budget", 0, -1, 16.0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetColumnWidth("Budget", 2, 1, 16.0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetColumnWidth("Budget", 0, 0, 256.0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetColumnWidth("Budget", 0, 0, Double.MIN_VALUE));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetColumnWidth("Budget", 0, 0, Double.POSITIVE_INFINITY));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetRowHeight("Budget", -1, 0, 28.5));
    assertThrows(
        NullPointerException.class, () -> new WorkbookCommand.SetRowHeight(null, 0, 0, 28.5));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCommand.SetRowHeight(" ", 0, 0, 28.5));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetRowHeight("Budget", 0, -1, 28.5));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetRowHeight("Budget", 2, 1, 28.5));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetRowHeight("Budget", 0, 0, (Short.MAX_VALUE / 20.0d) + 1.0d));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetRowHeight("Budget", 0, 0, Double.MIN_VALUE));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetRowHeight("Budget", 0, 0, 0.0));
    assertDoesNotThrow(() -> new WorkbookCommand.FreezePanes("Budget", 0, 2, 0, 2));
    assertDoesNotThrow(() -> new WorkbookCommand.FreezePanes("Budget", 2, 0, 2, 0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.FreezePanes("Budget", 0, 0, 0, 0));
    assertThrows(
        NullPointerException.class, () -> new WorkbookCommand.FreezePanes(null, 1, 1, 1, 1));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCommand.FreezePanes(" ", 1, 1, 1, 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.FreezePanes("Budget", -1, 1, 1, 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.FreezePanes("Budget", 1, -1, 1, 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.FreezePanes("Budget", 1, 1, -1, 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.FreezePanes("Budget", 1, 1, 1, -1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.FreezePanes("Budget", 0, 1, 1, 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.FreezePanes("Budget", 1, 0, 1, 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.FreezePanes("Budget", 2, 1, 1, 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.FreezePanes("Budget", 1, 2, 1, 1));
  }

  @Test
  void validatesCellRangeStyleAndAppendCommandInputs() {
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookCommand.SetCell(null, "A1", ExcelCellValue.text("x")));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookCommand.SetCell("Budget", null, ExcelCellValue.text("x")));
    assertThrows(
        NullPointerException.class, () -> new WorkbookCommand.SetCell("Budget", "A1", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetCell(" ", "A1", ExcelCellValue.text("x")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetCell("Budget", " ", ExcelCellValue.text("x")));
    assertThrows(
        NullPointerException.class, () -> new WorkbookCommand.SetRange(null, "A1", List.of()));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCommand.SetRange(" ", "A1", List.of()));
    assertThrows(
        NullPointerException.class, () -> new WorkbookCommand.SetRange("Budget", null, List.of()));
    assertThrows(
        NullPointerException.class, () -> new WorkbookCommand.SetRange("Budget", "A1", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetRange("Budget", " ", List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetRange("Budget", "A1", List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.SetRange("Budget", "A1:B2", List.of(List.of())));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookCommand.SetRange(
                "Budget",
                "A1:B2",
                List.of(
                    List.of(ExcelCellValue.text("x")),
                    List.of(ExcelCellValue.text("y"), ExcelCellValue.text("z")))));
    List<List<ExcelCellValue>> rowsWithNull = new ArrayList<>();
    List<ExcelCellValue> rowWithNull = new ArrayList<>();
    rowWithNull.add(null);
    rowsWithNull.add(rowWithNull);
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookCommand.SetRange("Budget", "A1", rowsWithNull));
    assertThrows(NullPointerException.class, () -> new WorkbookCommand.ClearRange(null, "A1"));
    assertThrows(NullPointerException.class, () -> new WorkbookCommand.ClearRange("Budget", null));
    assertThrows(IllegalArgumentException.class, () -> new WorkbookCommand.ClearRange(" ", "A1"));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCommand.ClearRange("Budget", " "));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookCommand.ApplyStyle(null, "A1", ExcelCellStyle.numberFormat("0")));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookCommand.ApplyStyle("Budget", null, ExcelCellStyle.numberFormat("0")));
    assertThrows(
        NullPointerException.class, () -> new WorkbookCommand.ApplyStyle("Budget", "A1", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.ApplyStyle(" ", "A1", ExcelCellStyle.numberFormat("0")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.ApplyStyle("Budget", " ", ExcelCellStyle.numberFormat("0")));
    assertThrows(NullPointerException.class, () -> new WorkbookCommand.AppendRow(null, List.of()));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCommand.AppendRow(" ", List.of()));
    assertThrows(NullPointerException.class, () -> new WorkbookCommand.AppendRow("Budget", null));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookCommand.AppendRow("Budget", List.of()));
    List<ExcelCellValue> valuesWithNull = new ArrayList<>();
    valuesWithNull.add(null);
    assertThrows(
        NullPointerException.class, () -> new WorkbookCommand.AppendRow("Budget", valuesWithNull));
    assertThrows(NullPointerException.class, () -> new WorkbookCommand.AutoSizeColumns(null));
    assertThrows(IllegalArgumentException.class, () -> new WorkbookCommand.AutoSizeColumns(" "));
  }
}
