package dev.erst.gridgrind.contract.action;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.NamedRangeScope;
import dev.erst.gridgrind.contract.dto.NamedRangeTarget;
import dev.erst.gridgrind.contract.dto.SheetCopyPosition;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import dev.erst.gridgrind.excel.foundation.ExcelColumnSpan;
import dev.erst.gridgrind.excel.foundation.ExcelRowSpan;
import java.util.ArrayList;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for action payload validation independent of step target selection. */
class MutationActionTest {
  @Test
  void exposesStableActionTypeNames() {
    assertEquals("ENSURE_SHEET", new WorkbookMutationAction.EnsureSheet().actionType());
    assertEquals("SET_CELL", new CellMutationAction.SetCell(textCell("Owner")).actionType());
    assertEquals(
        "CLEAR_WORKBOOK_PROTECTION",
        new WorkbookMutationAction.ClearWorkbookProtection().actionType());
  }

  @Test
  void copiesRectangularRowsAndValidatesAppendRows() {
    List<List<CellInput>> rows = new ArrayList<>();
    rows.add(new ArrayList<>(List.of(textCell("A"), new CellInput.Numeric(1.0d))));
    CellMutationAction.SetRange setRange = new CellMutationAction.SetRange(rows);

    rows.clear();

    assertEquals(1, setRange.rows().size());
    assertEquals(text("A"), ((CellInput.Text) setRange.rows().getFirst().getFirst()).source());
    assertEquals(
        2, new CellMutationAction.AppendRow(List.of(textCell("A"), textCell("B"))).values().size());
    assertThrows(IllegalArgumentException.class, () -> new CellMutationAction.AppendRow(List.of()));
  }

  @Test
  void validatesScalarActionArguments() {
    assertEquals(125, new WorkbookMutationAction.SetSheetZoom(125).zoomPercent());
    assertFalse(WorkbookMutationAction.GroupRows.expanded().collapsed());
    assertFalse(WorkbookMutationAction.GroupColumns.expanded().collapsed());
    assertEquals(
        new SheetCopyPosition.AppendAtEnd(),
        new WorkbookMutationAction.CopySheet("Budget Copy").position());
    assertEquals(
        "BudgetTotal",
        new StructuredMutationAction.SetNamedRange(
                "BudgetTotal", new NamedRangeScope.Workbook(), new NamedRangeTarget("Budget", "B4"))
            .name());

    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookMutationAction.SetSheetZoom(401));
    assertThrows(IllegalArgumentException.class, () -> new WorkbookMutationAction.RenameSheet(" "));
    assertThrows(IllegalArgumentException.class, () -> new WorkbookMutationAction.MoveSheet(-1));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new StructuredMutationAction.SetNamedRange(
                " ", new NamedRangeScope.Workbook(), new NamedRangeTarget("Budget", "B4")));

    assertThrows(
        NullPointerException.class,
        () -> new WorkbookMutationAction.CopySheet("Budget Copy", null));
    assertThrows(
        NullPointerException.class, () -> new WorkbookMutationAction.GroupRows((Boolean) null));
    assertThrows(
        NullPointerException.class, () -> new WorkbookMutationAction.GroupColumns((Boolean) null));
    assertThrows(
        NullPointerException.class, () -> new StructuredMutationAction.SetAutofilter(null, null));
  }

  @Test
  void validationHelpersRejectBoundsDuplicatesAndNonRectangularRows() {
    assertEquals(
        "field must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> MutationAction.Validation.requireNonBlank(" ", "field"))
            .getMessage());
    assertEquals(
        "index must not be negative",
        assertThrows(
                IllegalArgumentException.class,
                () -> MutationAction.Validation.requireNonNegative(-1, "index"))
            .getMessage());
    assertEquals(
        "count must be greater than 0",
        assertThrows(
                IllegalArgumentException.class,
                () -> MutationAction.Validation.requirePositive(0, "count"))
            .getMessage());
    assertEquals(
        "delta must not be 0",
        assertThrows(
                IllegalArgumentException.class,
                () -> MutationAction.Validation.requireNonZero(0, "delta"))
            .getMessage());
    assertEquals(
        "row must not exceed " + ExcelRowSpan.MAX_ROW_INDEX + " (Excel row limit)",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    MutationAction.Validation.requireRowIndex(
                        ExcelRowSpan.MAX_ROW_INDEX + 1, "row"))
            .getMessage());
    assertEquals(
        "column must not exceed " + ExcelColumnSpan.MAX_COLUMN_INDEX + " (Excel column limit)",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    MutationAction.Validation.requireColumnIndex(
                        ExcelColumnSpan.MAX_COLUMN_INDEX + 1, "column"))
            .getMessage());
    assertEquals(
        "last must not be less than first",
        assertThrows(
                IllegalArgumentException.class,
                () -> MutationAction.Validation.requireOrderedSpan(2, 1, "first", "last"))
            .getMessage());
    assertEquals(
        "widthCharacters must not exceed 255.0 (Excel column width limit): got 255.1",
        assertThrows(
                IllegalArgumentException.class,
                () -> MutationAction.Validation.requireColumnWidthCharacters(255.1d))
            .getMessage());
    assertEquals(
        "widthCharacters is too small to produce a visible Excel column width: got 1.0E-4",
        assertThrows(
                IllegalArgumentException.class,
                () -> MutationAction.Validation.requireColumnWidthCharacters(0.0001d))
            .getMessage());
    assertEquals(
        "heightPoints must not exceed 409.0 (Excel row height limit): got 2000.0",
        assertThrows(
                IllegalArgumentException.class,
                () -> MutationAction.Validation.requireRowHeightPoints(2000.0d))
            .getMessage());
    assertEquals(
        "heightPoints is too small to produce a visible Excel row height: 0.01",
        assertThrows(
                IllegalArgumentException.class,
                () -> MutationAction.Validation.requireRowHeightPoints(0.01d))
            .getMessage());
    assertEquals(
        "widthCharacters must be finite",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    MutationAction.Validation.requireColumnWidthCharacters(
                        Double.POSITIVE_INFINITY))
            .getMessage());
    assertEquals(
        "heightPoints must be greater than 0",
        assertThrows(
                IllegalArgumentException.class,
                () -> MutationAction.Validation.requireRowHeightPoints(0.0d))
            .getMessage());
    assertEquals(
        "zoomPercent must be between 10 and 400 inclusive: 9",
        assertThrows(
                IllegalArgumentException.class,
                () -> MutationAction.Validation.requireZoomPercent(9))
            .getMessage());
    assertEquals(
        "zoomPercent must be between 10 and 400 inclusive: 401",
        assertThrows(
                IllegalArgumentException.class,
                () -> MutationAction.Validation.requireZoomPercent(401))
            .getMessage());
    assertEquals(
        "rows must not be null",
        assertThrows(NullPointerException.class, () -> MutationAction.Validation.copyRows(null))
            .getMessage());
    assertEquals(
        java.util.Arrays.asList((List<CellInput>) null),
        MutationAction.Validation.copyRows(java.util.Arrays.asList((List<CellInput>) null)));
    MutationAction.Validation.requireDistinct(List.of("a", "b"), "sheetNames");
    assertEquals(
        List.of("Budget", "Ops"),
        MutationAction.Validation.copySheetNames(List.of("Budget", "Ops"), "sheetNames"));
    assertEquals(
        "name must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> MutationAction.Validation.requirePivotTableName(" "))
            .getMessage());
    MutationAction.Validation.requirePivotTableName("1bad");
    assertEquals(
        "sheetNames must not contain duplicates",
        assertThrows(
                IllegalArgumentException.class,
                () -> MutationAction.Validation.requireDistinct(List.of("a", "a"), "sheetNames"))
            .getMessage());
    assertEquals(
        "rows must not contain null rows",
        assertThrows(
                NullPointerException.class,
                () ->
                    MutationAction.Validation.requireRectangularRows(
                        java.util.Arrays.asList((List<CellInput>) null)))
            .getMessage());
    assertEquals(
        "rows must not contain empty rows",
        assertThrows(
                IllegalArgumentException.class,
                () -> MutationAction.Validation.requireRectangularRows(List.of(List.of())))
            .getMessage());
    assertEquals(
        "rows must describe a rectangular matrix",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    MutationAction.Validation.requireRectangularRows(
                        List.of(List.of(textCell("A")), List.of(textCell("A"), textCell("B")))))
            .getMessage());
    assertEquals(
        "rows must not contain null cell values",
        assertThrows(
                NullPointerException.class,
                () ->
                    MutationAction.Validation.requireRectangularRows(
                        List.of(java.util.Arrays.asList(textCell("A"), null))))
            .getMessage());
    assertEquals(
        "rows must not be empty",
        assertThrows(
                IllegalArgumentException.class,
                () -> MutationAction.Validation.requireRectangularRows(List.of()))
            .getMessage());
    assertEquals(
        List.of(List.of(textCell("A"))),
        MutationAction.Validation.freezeRows(
            MutationAction.Validation.copyRows(List.of(List.of(textCell("A"))))));
    MutationAction.Validation.requireNonBlank("ok", "field");
    MutationAction.Validation.requirePositive(1, "count");
    MutationAction.Validation.requireRowIndex(ExcelRowSpan.MAX_ROW_INDEX, "row");
    MutationAction.Validation.requireColumnIndex(ExcelColumnSpan.MAX_COLUMN_INDEX, "column");
    MutationAction.Validation.requireOrderedSpan(1, 1, "first", "last");
  }

  @Test
  void actionConstructorsCoverNullCollectionDefaulting() {
    assertEquals(List.of(), new StructuredMutationAction.SetAutofilter().criteria());
    assertEquals(
        "criteria must not contain null values",
        assertThrows(
                NullPointerException.class,
                () ->
                    new StructuredMutationAction.SetAutofilter(
                        java.util.Arrays.asList(
                            (dev.erst.gridgrind.contract.dto.AutofilterFilterColumnInput) null),
                        null))
            .getMessage());
    assertEquals(
        "values must not be null",
        assertThrows(NullPointerException.class, () -> new CellMutationAction.AppendRow(null))
            .getMessage());
  }

  private static CellInput.Text textCell(String value) {
    return new CellInput.Text(text(value));
  }

  private static TextSourceInput text(String value) {
    return TextSourceInput.inline(value);
  }
}
