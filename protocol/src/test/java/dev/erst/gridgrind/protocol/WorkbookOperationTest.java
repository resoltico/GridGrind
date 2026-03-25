package dev.erst.gridgrind.protocol;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.ExcelHorizontalAlignment;
import dev.erst.gridgrind.excel.ExcelVerticalAlignment;
import java.util.ArrayList;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for WorkbookOperation record construction and operationType behavior. */
class WorkbookOperationTest {
  @Test
  void buildsSupportedOperationsAndCopiesCollections() {
    List<CellInput> rowValues = new ArrayList<>(List.of(new CellInput.Text("Item")));
    List<List<CellInput>> rows =
        new ArrayList<>(
            List.of(
                new ArrayList<>(List.of(new CellInput.Text("Item"), new CellInput.Numeric(12.0))),
                new ArrayList<>(List.of(new CellInput.Text("Tax"), new CellInput.Numeric(3.0)))));
    CellStyleInput style =
        new CellStyleInput(
            "#,##0.00",
            true,
            null,
            true,
            ExcelHorizontalAlignment.RIGHT,
            ExcelVerticalAlignment.CENTER);

    WorkbookOperation.EnsureSheet ensureSheet = new WorkbookOperation.EnsureSheet("Budget");
    WorkbookOperation.RenameSheet renameSheet =
        new WorkbookOperation.RenameSheet("Budget", "Summary");
    WorkbookOperation.DeleteSheet deleteSheet = new WorkbookOperation.DeleteSheet("Archive");
    WorkbookOperation.MoveSheet moveSheet = new WorkbookOperation.MoveSheet("Budget", 1);
    WorkbookOperation.MergeCells mergeCells = new WorkbookOperation.MergeCells("Budget", "A1:B2");
    WorkbookOperation.UnmergeCells unmergeCells =
        new WorkbookOperation.UnmergeCells("Budget", "A1:B2");
    WorkbookOperation.SetColumnWidth setColumnWidth =
        new WorkbookOperation.SetColumnWidth("Budget", 0, 2, 16.0);
    WorkbookOperation.SetRowHeight setRowHeight =
        new WorkbookOperation.SetRowHeight("Budget", 0, 3, 28.5);
    WorkbookOperation.FreezePanes freezePanes =
        new WorkbookOperation.FreezePanes("Budget", 1, 2, 1, 2);
    WorkbookOperation.SetCell setCell =
        new WorkbookOperation.SetCell("Budget", "A1", new CellInput.Text("Item"));
    WorkbookOperation.SetRange setRange = new WorkbookOperation.SetRange("Budget", "A1:B2", rows);
    WorkbookOperation.ClearRange clearRange = new WorkbookOperation.ClearRange("Budget", "C1:C4");
    WorkbookOperation.ApplyStyle applyStyle =
        new WorkbookOperation.ApplyStyle("Budget", "B1:B2", style);
    WorkbookOperation.AppendRow appendRow = new WorkbookOperation.AppendRow("Budget", rowValues);
    WorkbookOperation.AutoSizeColumns autoSizeColumns =
        new WorkbookOperation.AutoSizeColumns("Budget");
    WorkbookOperation.EvaluateFormulas evaluateFormulas = new WorkbookOperation.EvaluateFormulas();
    WorkbookOperation.ForceFormulaRecalculationOnOpen recalcOnOpen =
        new WorkbookOperation.ForceFormulaRecalculationOnOpen();

    rowValues.clear();
    rows.clear();

    assertEquals("Budget", ensureSheet.sheetName());
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
    assertEquals("C1:C4", clearRange.range());
    assertEquals(style, applyStyle.style());
    assertEquals(1, appendRow.values().size());
    assertEquals("Budget", autoSizeColumns.sheetName());
    assertEquals("EVALUATE_FORMULAS", evaluateFormulas.operationType());
    assertEquals("FORCE_FORMULA_RECALCULATION_ON_OPEN", recalcOnOpen.operationType());
  }

  @Test
  void validatesNullAndEmptyCollectionConstraints() {
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookOperation.AppendRow("Budget", null));
    assertThrows(NullPointerException.class, () -> new WorkbookOperation.AutoSizeColumns(null));
    assertThrows(IllegalArgumentException.class, () -> new WorkbookOperation.AutoSizeColumns(" "));
    assertThrows(
        NullPointerException.class, () -> new WorkbookOperation.RenameSheet(null, "Summary"));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookOperation.RenameSheet(" ", "Summary"));
    assertThrows(
        NullPointerException.class, () -> new WorkbookOperation.RenameSheet("Budget", null));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookOperation.RenameSheet("Budget", " "));
    assertThrows(NullPointerException.class, () -> new WorkbookOperation.DeleteSheet(null));
    assertThrows(IllegalArgumentException.class, () -> new WorkbookOperation.DeleteSheet(" "));
    assertThrows(NullPointerException.class, () -> new WorkbookOperation.MoveSheet("Budget", null));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookOperation.MoveSheet("Budget", -1));
    assertThrows(NullPointerException.class, () -> new WorkbookOperation.MergeCells(null, "A1:B2"));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookOperation.MergeCells("Budget", " "));
    assertThrows(
        NullPointerException.class, () -> new WorkbookOperation.UnmergeCells("Budget", null));
    assertThrows(
        NullPointerException.class, () -> new WorkbookOperation.SetColumnWidth(null, 0, 0, 16.0));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookOperation.SetColumnWidth("Budget", null, 0, 16.0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookOperation.SetColumnWidth("Budget", 1, 0, 16.0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookOperation.SetColumnWidth("Budget", 0, 0, 0.0));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookOperation.SetRowHeight("Budget", null, 0, 28.5));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookOperation.SetRowHeight("Budget", 2, 1, 28.5));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookOperation.SetRowHeight("Budget", 0, 0, Double.NaN));
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookOperation.FreezePanes("Budget", null, 1, 0, 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookOperation.FreezePanes("Budget", 0, 0, 0, 0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookOperation.FreezePanes("Budget", 0, 1, 1, 1));
  }

  @Test
  void validatesOperationRequirements() {
    assertThrows(IllegalArgumentException.class, () -> new WorkbookOperation.EnsureSheet(" "));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookOperation.RenameSheet("Budget", " "));
    assertThrows(IllegalArgumentException.class, () -> new WorkbookOperation.DeleteSheet(" "));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookOperation.MoveSheet("Budget", -1));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookOperation.MergeCells("Budget", " "));
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookOperation.UnmergeCells("Budget", " "));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookOperation.SetColumnWidth("Budget", -1, 0, 16.0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookOperation.SetRowHeight("Budget", 0, -1, 28.5));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookOperation.FreezePanes("Budget", 1, 0, 0, 0));

    assertThrows(
        NullPointerException.class,
        () -> new WorkbookOperation.SetCell("Budget", null, new CellInput.Text("x")));
    assertThrows(
        NullPointerException.class, () -> new WorkbookOperation.SetCell("Budget", "A1", null));

    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookOperation.SetRange("Budget", "A1:B2", List.of()));

    List<List<CellInput>> rowsWithNullRow = new ArrayList<>();
    rowsWithNullRow.add(null);
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookOperation.SetRange("Budget", "A1:B1", rowsWithNullRow));

    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookOperation.SetRange("Budget", "A1:B2", List.of(List.of())));

    assertThrows(
        IllegalArgumentException.class,
        () ->
            new WorkbookOperation.SetRange(
                "Budget",
                "A1:B2",
                List.of(
                    List.of(new CellInput.Text("x")),
                    List.of(new CellInput.Text("y"), new CellInput.Text("z")))));

    List<List<CellInput>> rowsWithNullValue = new ArrayList<>();
    List<CellInput> rowWithNullValue = new ArrayList<>();
    rowWithNullValue.add(null);
    rowsWithNullValue.add(rowWithNullValue);
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookOperation.SetRange("Budget", "A1", rowsWithNullValue));

    assertThrows(
        NullPointerException.class,
        () -> new WorkbookOperation.ApplyStyle("Budget", "A1:A2", null));

    List<CellInput> valuesWithNull = new ArrayList<>();
    valuesWithNull.add(null);
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookOperation.AppendRow("Budget", valuesWithNull));

    // null rows list is coalesced to empty, which then fails the non-empty check
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookOperation.SetRange("Budget", "A1", null));
  }

  @Test
  void validatesColumnWidthHelperBranches() {
    assertDoesNotThrow(() -> WorkbookOperation.Validation.requireColumnWidthCharacters(8.43d));
    assertThrows(
        IllegalArgumentException.class,
        () -> WorkbookOperation.Validation.requireColumnWidthCharacters(256.0));
    assertThrows(
        IllegalArgumentException.class,
        () -> WorkbookOperation.Validation.requireColumnWidthCharacters(Double.MIN_VALUE));
  }

  @Test
  void validatesRowHeightHelperBranches() {
    assertDoesNotThrow(() -> WorkbookOperation.Validation.requireRowHeightPoints(15.0d));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            WorkbookOperation.Validation.requireRowHeightPoints((Short.MAX_VALUE / 20.0d) + 1.0d));
    assertThrows(
        IllegalArgumentException.class,
        () -> WorkbookOperation.Validation.requireRowHeightPoints(Double.MIN_VALUE));
  }

  @Test
  void validatesFreezePaneCoordinateHelperBranches() {
    assertDoesNotThrow(() -> WorkbookOperation.Validation.requireFreezePaneCoordinates(1, 2, 1, 2));
    assertDoesNotThrow(() -> WorkbookOperation.Validation.requireFreezePaneCoordinates(0, 2, 0, 2));
    assertDoesNotThrow(() -> WorkbookOperation.Validation.requireFreezePaneCoordinates(2, 0, 2, 0));
    assertThrows(
        IllegalArgumentException.class,
        () -> WorkbookOperation.Validation.requireFreezePaneCoordinates(0, 0, 0, 0));
    assertThrows(
        IllegalArgumentException.class,
        () -> WorkbookOperation.Validation.requireFreezePaneCoordinates(0, 1, 1, 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> WorkbookOperation.Validation.requireFreezePaneCoordinates(1, 0, 1, 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> WorkbookOperation.Validation.requireFreezePaneCoordinates(2, 1, 1, 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> WorkbookOperation.Validation.requireFreezePaneCoordinates(1, 2, 1, 1));
  }

  @Test
  void operationTypeCoversAllSubtypes() {
    CellInput textValue = new CellInput.Text("x");
    CellStyleInput style = new CellStyleInput(null, false, null, false, null, null);

    assertEquals("ENSURE_SHEET", new WorkbookOperation.EnsureSheet("Budget").operationType());
    assertEquals(
        "RENAME_SHEET", new WorkbookOperation.RenameSheet("Budget", "Summary").operationType());
    assertEquals("DELETE_SHEET", new WorkbookOperation.DeleteSheet("Budget").operationType());
    assertEquals("MOVE_SHEET", new WorkbookOperation.MoveSheet("Budget", 0).operationType());
    assertEquals(
        "MERGE_CELLS", new WorkbookOperation.MergeCells("Budget", "A1:B2").operationType());
    assertEquals(
        "UNMERGE_CELLS", new WorkbookOperation.UnmergeCells("Budget", "A1:B2").operationType());
    assertEquals(
        "SET_COLUMN_WIDTH",
        new WorkbookOperation.SetColumnWidth("Budget", 0, 1, 16.0).operationType());
    assertEquals(
        "SET_ROW_HEIGHT", new WorkbookOperation.SetRowHeight("Budget", 0, 1, 28.5).operationType());
    assertEquals(
        "FREEZE_PANES", new WorkbookOperation.FreezePanes("Budget", 1, 2, 1, 2).operationType());
    assertEquals(
        "SET_CELL", new WorkbookOperation.SetCell("Budget", "A1", textValue).operationType());
    assertEquals(
        "SET_RANGE",
        new WorkbookOperation.SetRange("Budget", "A1", List.of(List.of(textValue)))
            .operationType());
    assertEquals("CLEAR_RANGE", new WorkbookOperation.ClearRange("Budget", "A1").operationType());
    assertEquals(
        "APPLY_STYLE", new WorkbookOperation.ApplyStyle("Budget", "A1", style).operationType());
    assertEquals(
        "APPEND_ROW",
        new WorkbookOperation.AppendRow("Budget", List.of(textValue)).operationType());
    assertEquals(
        "AUTO_SIZE_COLUMNS", new WorkbookOperation.AutoSizeColumns("Budget").operationType());
    assertEquals("EVALUATE_FORMULAS", new WorkbookOperation.EvaluateFormulas().operationType());
    assertEquals(
        "FORCE_FORMULA_RECALCULATION_ON_OPEN",
        new WorkbookOperation.ForceFormulaRecalculationOnOpen().operationType());
  }
}
