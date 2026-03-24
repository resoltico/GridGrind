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
    List<String> columns = new ArrayList<>(List.of("A"));
    ExcelCellStyle style =
        new ExcelCellStyle(
            "#,##0.00",
            true,
            false,
            true,
            ExcelHorizontalAlignment.RIGHT,
            ExcelVerticalAlignment.CENTER);

    WorkbookCommand.CreateSheet createSheet = new WorkbookCommand.CreateSheet("Budget");
    WorkbookCommand.SetCell setCell =
        new WorkbookCommand.SetCell("Budget", "A1", ExcelCellValue.date(LocalDate.of(2026, 3, 23)));
    WorkbookCommand.SetRange setRange = new WorkbookCommand.SetRange("Budget", "A1:B2", rows);
    WorkbookCommand.ClearRange clearRange = new WorkbookCommand.ClearRange("Budget", "C1:C2");
    WorkbookCommand.ApplyStyle applyStyle =
        new WorkbookCommand.ApplyStyle("Budget", "A1:B1", style);
    WorkbookCommand.AppendRow appendRow = new WorkbookCommand.AppendRow("Budget", values);
    WorkbookCommand.AutoSizeColumns autoSizeColumns =
        new WorkbookCommand.AutoSizeColumns("Budget", columns);
    WorkbookCommand.EvaluateAllFormulas evaluate = new WorkbookCommand.EvaluateAllFormulas();
    WorkbookCommand.ForceFormulaRecalculationOnOpen recalc =
        new WorkbookCommand.ForceFormulaRecalculationOnOpen();

    values.clear();
    rows.clear();
    columns.clear();

    assertEquals("Budget", createSheet.sheetName());
    assertEquals("A1", setCell.address());
    assertEquals("A1:B2", setRange.range());
    assertEquals(2, setRange.rows().size());
    assertEquals("C1:C2", clearRange.range());
    assertEquals(style, applyStyle.style());
    assertEquals(1, appendRow.values().size());
    assertEquals(List.of("A"), autoSizeColumns.columnNames());
    assertNotNull(evaluate);
    assertNotNull(recalc);
  }

  @Test
  void validatesCommandInputs() {
    assertThrows(NullPointerException.class, () -> new WorkbookCommand.CreateSheet(null));
    assertThrows(IllegalArgumentException.class, () -> new WorkbookCommand.CreateSheet(" "));
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
    assertThrows(
        NullPointerException.class, () -> new WorkbookCommand.AutoSizeColumns(null, List.of("A")));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.AutoSizeColumns(" ", List.of("A")));
    assertThrows(
        NullPointerException.class, () -> new WorkbookCommand.AutoSizeColumns("Budget", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.AutoSizeColumns("Budget", List.of()));
    List<String> columnsWithNull = new ArrayList<>();
    columnsWithNull.add(null);
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookCommand.AutoSizeColumns("Budget", columnsWithNull));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookCommand.AutoSizeColumns("Budget", List.of(" ")));
  }
}
