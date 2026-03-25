package dev.erst.gridgrind.protocol;

import static org.junit.jupiter.api.Assertions.*;

import java.util.ArrayList;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for WorkbookOperation record construction and default method behavior. */
class WorkbookOperationTest {
  @Test
  void buildsSupportedOperationsAndCopiesCollections() {
    List<CellInput> rowValues = new ArrayList<>(List.of(new CellInput.Text("Item")));
    List<List<CellInput>> rows =
        new ArrayList<>(
            List.of(
                new ArrayList<>(List.of(new CellInput.Text("Item"), new CellInput.Numeric(12.0))),
                new ArrayList<>(List.of(new CellInput.Text("Tax"), new CellInput.Numeric(3.0)))));
    List<String> columns = new ArrayList<>(List.of("A"));
    CellStyleInput style =
        new CellStyleInput(
            "#,##0.00",
            true,
            null,
            true,
            CellStyleInput.HorizontalAlignmentInput.RIGHT,
            CellStyleInput.VerticalAlignmentInput.CENTER);

    WorkbookOperation.EnsureSheet ensureSheet = new WorkbookOperation.EnsureSheet("Budget");
    WorkbookOperation.SetCell setCell =
        new WorkbookOperation.SetCell("Budget", "A1", new CellInput.Text("Item"));
    WorkbookOperation.SetRange setRange = new WorkbookOperation.SetRange("Budget", "A1:B2", rows);
    WorkbookOperation.ClearRange clearRange = new WorkbookOperation.ClearRange("Budget", "C1:C4");
    WorkbookOperation.ApplyStyle applyStyle =
        new WorkbookOperation.ApplyStyle("Budget", "B1:B2", style);
    WorkbookOperation.AppendRow appendRow = new WorkbookOperation.AppendRow("Budget", rowValues);
    WorkbookOperation.AutoSizeColumns autoSizeColumns =
        new WorkbookOperation.AutoSizeColumns("Budget", columns);
    WorkbookOperation.EvaluateFormulas evaluateFormulas = new WorkbookOperation.EvaluateFormulas();
    WorkbookOperation.ForceFormulaRecalculationOnOpen recalcOnOpen =
        new WorkbookOperation.ForceFormulaRecalculationOnOpen();

    rowValues.clear();
    rows.clear();
    columns.clear();

    assertEquals("ENSURE_SHEET", ensureSheet.operationType());
    assertEquals("A1", setCell.address());
    assertEquals("A1:B2", setRange.range());
    assertEquals(2, setRange.rows().size());
    assertEquals("C1:C4", clearRange.range());
    assertEquals(style, applyStyle.style());
    assertEquals(1, appendRow.values().size());
    assertEquals(List.of("A"), autoSizeColumns.columns());
    assertEquals("EVALUATE_FORMULAS", evaluateFormulas.operationType());
    assertEquals("FORCE_FORMULA_RECALCULATION_ON_OPEN", recalcOnOpen.operationType());
  }

  @Test
  void validatesNullAndEmptyCollectionConstraints() {
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookOperation.AppendRow("Budget", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookOperation.AutoSizeColumns("Budget", null));
  }

  @Test
  void validatesOperationRequirements() {
    assertThrows(IllegalArgumentException.class, () -> new WorkbookOperation.EnsureSheet(" "));

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

    List<String> columnsWithNull = new ArrayList<>();
    columnsWithNull.add(null);
    assertThrows(
        NullPointerException.class,
        () -> new WorkbookOperation.AutoSizeColumns("Budget", columnsWithNull));

    assertThrows(
        IllegalArgumentException.class,
        () -> new WorkbookOperation.AutoSizeColumns("Budget", List.of(" ")));

    // null rows list is coalesced to empty, which then fails the non-empty check
    assertThrows(
        IllegalArgumentException.class, () -> new WorkbookOperation.SetRange("Budget", "A1", null));
  }

  @Test
  void defaultMethodsReturnCorrectValuesForAllSubtypes() {
    CellStyleInput style = new CellStyleInput(null, false, null, false, null, null);
    CellInput textValue = new CellInput.Text("x");

    WorkbookOperation.EnsureSheet ensureSheet = new WorkbookOperation.EnsureSheet("Budget");
    WorkbookOperation.SetCell setCell = new WorkbookOperation.SetCell("Budget", "A1", textValue);
    WorkbookOperation.SetCell setCellWithFormula =
        new WorkbookOperation.SetCell("Budget", "A1", new CellInput.Formula("SUM(B1:B2)"));
    WorkbookOperation.SetRange setRange =
        new WorkbookOperation.SetRange("Budget", "A1:B2", List.of(List.of(textValue, textValue)));
    WorkbookOperation.SetRange setRangeWithAddress =
        new WorkbookOperation.SetRange("Budget", "A1", List.of(List.of(textValue)));
    WorkbookOperation.ClearRange clearRange = new WorkbookOperation.ClearRange("Budget", "C1:C2");
    WorkbookOperation.ApplyStyle applyStyle =
        new WorkbookOperation.ApplyStyle("Budget", "D1:D2", style);
    WorkbookOperation.AppendRow appendRow =
        new WorkbookOperation.AppendRow("Budget", List.of(textValue));
    WorkbookOperation.AutoSizeColumns autoSizeColumns =
        new WorkbookOperation.AutoSizeColumns("Budget", List.of("A"));
    WorkbookOperation.EvaluateFormulas evaluateFormulas = new WorkbookOperation.EvaluateFormulas();
    WorkbookOperation.ForceFormulaRecalculationOnOpen recalcOnOpen =
        new WorkbookOperation.ForceFormulaRecalculationOnOpen();

    // operationType for all subtypes
    assertEquals("SET_CELL", setCell.operationType());
    assertEquals("SET_RANGE", setRange.operationType());
    assertEquals("CLEAR_RANGE", clearRange.operationType());
    assertEquals("APPLY_STYLE", applyStyle.operationType());
    assertEquals("APPEND_ROW", appendRow.operationType());
    assertEquals("AUTO_SIZE_COLUMNS", autoSizeColumns.operationType());

    // extractSheetName — subtypes with a sheet
    assertEquals("Budget", ensureSheet.extractSheetName());
    assertEquals("Budget", setCell.extractSheetName());
    assertEquals("Budget", setRange.extractSheetName());
    assertEquals("Budget", clearRange.extractSheetName());
    assertEquals("Budget", applyStyle.extractSheetName());
    assertEquals("Budget", appendRow.extractSheetName());
    assertEquals("Budget", autoSizeColumns.extractSheetName());
    // extractSheetName — subtypes without a sheet
    assertNull(evaluateFormulas.extractSheetName());
    assertNull(recalcOnOpen.extractSheetName());

    // extractAddress — only SetCell returns non-null
    assertEquals("A1", setCell.extractAddress());
    assertNull(ensureSheet.extractAddress());
    assertNull(setRange.extractAddress());
    assertNull(clearRange.extractAddress());
    assertNull(applyStyle.extractAddress());
    assertNull(appendRow.extractAddress());
    assertNull(autoSizeColumns.extractAddress());
    assertNull(evaluateFormulas.extractAddress());
    assertNull(recalcOnOpen.extractAddress());

    // extractRange — SetRange, ClearRange, ApplyStyle return non-null
    assertEquals("A1:B2", setRange.extractRange());
    assertEquals("C1:C2", clearRange.extractRange());
    assertEquals("D1:D2", applyStyle.extractRange());
    assertNull(ensureSheet.extractRange());
    assertNull(setCell.extractRange());
    assertNull(appendRow.extractRange());
    assertNull(autoSizeColumns.extractRange());
    assertNull(evaluateFormulas.extractRange());
    assertNull(recalcOnOpen.extractRange());

    // extractValue — only SetCell returns non-null
    assertEquals(textValue, setCell.extractValue());
    assertInstanceOf(CellInput.Formula.class, setCellWithFormula.extractValue());
    assertNull(ensureSheet.extractValue());
    assertNull(setRangeWithAddress.extractValue());
    assertNull(clearRange.extractValue());
    assertNull(applyStyle.extractValue());
    assertNull(appendRow.extractValue());
    assertNull(autoSizeColumns.extractValue());
    assertNull(evaluateFormulas.extractValue());
    assertNull(recalcOnOpen.extractValue());
  }
}
