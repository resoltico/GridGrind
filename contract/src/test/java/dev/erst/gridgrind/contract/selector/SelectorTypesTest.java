package dev.erst.gridgrind.contract.selector;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.RichTextRunInput;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Validation coverage for selector-first protocol targeting types. */
class SelectorTypesTest {
  @Test
  void allSelectorFamiliesExposeExpectedCardinalityContracts() {
    assertEquals(SelectorCardinality.EXACTLY_ONE, new WorkbookSelector.Current().cardinality());
    assertEquals(SelectorCardinality.ANY_NUMBER, new SheetSelector.All().cardinality());
    assertEquals(SelectorCardinality.EXACTLY_ONE, new SheetSelector.ByName("Budget").cardinality());
    assertEquals(
        SelectorCardinality.ONE_OR_MORE,
        new SheetSelector.ByNames(List.of("Budget", "Ops")).cardinality());

    assertEquals(
        SelectorCardinality.ANY_NUMBER, new CellSelector.AllUsedInSheet("Budget").cardinality());
    assertEquals(
        SelectorCardinality.EXACTLY_ONE, new CellSelector.ByAddress("Budget", "A1").cardinality());
    assertEquals(
        SelectorCardinality.ONE_OR_MORE,
        new CellSelector.ByAddresses("Budget", List.of("A1", "B2")).cardinality());
    assertEquals(
        SelectorCardinality.ONE_OR_MORE,
        new CellSelector.ByQualifiedAddresses(
                List.of(
                    new CellSelector.QualifiedAddress("Budget", "A1"),
                    new CellSelector.QualifiedAddress("Ops", "B2")))
            .cardinality());

    assertEquals(
        SelectorCardinality.ANY_NUMBER, new RangeSelector.AllOnSheet("Budget").cardinality());
    assertEquals(
        SelectorCardinality.EXACTLY_ONE,
        new RangeSelector.ByRange("Budget", "A1:B2").cardinality());
    assertEquals(
        SelectorCardinality.ONE_OR_MORE,
        new RangeSelector.ByRanges("Budget", List.of("A1:A2", "C3:D4")).cardinality());
    assertEquals(
        SelectorCardinality.EXACTLY_ONE,
        new RangeSelector.RectangularWindow("Budget", "B3", 2, 3).cardinality());

    assertEquals(
        SelectorCardinality.EXACTLY_ONE, new RowBandSelector.Span("Budget", 0, 2).cardinality());
    assertEquals(
        SelectorCardinality.EXACTLY_ONE,
        new RowBandSelector.Insertion("Budget", 5, 2).cardinality());
    assertEquals(
        SelectorCardinality.EXACTLY_ONE, new ColumnBandSelector.Span("Budget", 1, 3).cardinality());
    assertEquals(
        SelectorCardinality.EXACTLY_ONE,
        new ColumnBandSelector.Insertion("Budget", 4, 1).cardinality());

    assertEquals(SelectorCardinality.ANY_NUMBER, new TableSelector.All().cardinality());
    assertEquals(
        SelectorCardinality.EXACTLY_ONE, new TableSelector.ByName("BudgetTable").cardinality());
    assertEquals(
        SelectorCardinality.ONE_OR_MORE,
        new TableSelector.ByNames(List.of("BudgetTable", "OpsTable")).cardinality());
    assertEquals(
        SelectorCardinality.EXACTLY_ONE,
        new TableSelector.ByNameOnSheet("BudgetTable", "Budget").cardinality());

    TableSelector.ByName table = new TableSelector.ByName("BudgetTable");
    assertEquals(SelectorCardinality.ANY_NUMBER, new TableRowSelector.AllRows(table).cardinality());
    assertEquals(
        SelectorCardinality.EXACTLY_ONE, new TableRowSelector.ByIndex(table, 2).cardinality());
    assertEquals(
        SelectorCardinality.ZERO_OR_ONE,
        new TableRowSelector.ByKeyCell(table, "Item", new CellInput.Text(text("Laptop")))
            .cardinality());
    assertEquals(
        SelectorCardinality.ZERO_OR_ONE,
        new TableCellSelector.ByColumnName(new TableRowSelector.ByIndex(table, 2), "Amount")
            .cardinality());

    assertEquals(SelectorCardinality.ANY_NUMBER, new NamedRangeSelector.All().cardinality());
    assertEquals(
        SelectorCardinality.ONE_OR_MORE,
        new NamedRangeSelector.ByName("BudgetTotal").cardinality());
    assertEquals(
        SelectorCardinality.ONE_OR_MORE,
        new NamedRangeSelector.ByNames(List.of("BudgetTotal", "OpsTotal")).cardinality());
    assertEquals(
        SelectorCardinality.EXACTLY_ONE,
        new NamedRangeSelector.WorkbookScope("BudgetTotal").cardinality());
    assertEquals(
        SelectorCardinality.EXACTLY_ONE,
        new NamedRangeSelector.SheetScope("LocalItem", "Budget").cardinality());
    assertEquals(
        SelectorCardinality.ONE_OR_MORE,
        new NamedRangeSelector.AnyOf(
                List.of(
                    new NamedRangeSelector.WorkbookScope("BudgetTotal"),
                    new NamedRangeSelector.SheetScope("LocalItem", "Budget")))
            .cardinality());

    assertEquals(
        SelectorCardinality.ANY_NUMBER,
        new DrawingObjectSelector.AllOnSheet("Budget").cardinality());
    assertEquals(
        SelectorCardinality.EXACTLY_ONE,
        new DrawingObjectSelector.ByName("Budget", "OpsPicture").cardinality());
    assertEquals(
        SelectorCardinality.ANY_NUMBER, new ChartSelector.AllOnSheet("Budget").cardinality());
    assertEquals(
        SelectorCardinality.EXACTLY_ONE,
        new ChartSelector.ByName("Budget", "OpsChart").cardinality());
    assertEquals(SelectorCardinality.ANY_NUMBER, new PivotTableSelector.All().cardinality());
    assertEquals(
        SelectorCardinality.EXACTLY_ONE,
        new PivotTableSelector.ByName("Sales Pivot 2026").cardinality());
    assertEquals(
        SelectorCardinality.ONE_OR_MORE,
        new PivotTableSelector.ByNames(List.of("Sales Pivot 2026", "Ops Pivot")).cardinality());
    assertEquals(
        SelectorCardinality.EXACTLY_ONE,
        new PivotTableSelector.ByNameOnSheet("Sales Pivot 2026", "Report").cardinality());
  }

  @Test
  void sheetCellAndRangeSelectorsValidateAndNormalize() {
    assertEquals(new SheetSelector.All(), new SheetSelector.All());
    assertEquals(new SheetSelector.ByName("Budget"), new SheetSelector.ByName("Budget"));
    assertEquals(
        List.of("Budget", "Ops"), new SheetSelector.ByNames(List.of("Budget", "Ops")).names());
    assertThrows(IllegalArgumentException.class, () -> new SheetSelector.ByNames(List.of()));

    CellSelector.ByAddresses cells = new CellSelector.ByAddresses("Budget", List.of("A1", "$B$2"));
    assertEquals("Budget", cells.sheetName());
    assertEquals(List.of("A1", "$B$2"), cells.addresses());
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellSelector.ByAddresses("Budget", List.of("A1", "a1")));
    assertEquals(
        List.of(new CellSelector.QualifiedAddress("Budget", "A1")),
        new CellSelector.ByQualifiedAddresses(
                List.of(new CellSelector.QualifiedAddress("Budget", "A1")))
            .cells());
    assertEquals("Budget!A1", new CellSelector.QualifiedAddress("Budget", "A1").toString());

    RangeSelector.RectangularWindow window =
        new RangeSelector.RectangularWindow("Budget", "B3", 2, 3);
    assertEquals("B3:D4", window.range());
    assertThrows(
        IllegalArgumentException.class,
        () -> new RangeSelector.ByRanges("Budget", List.of("A1:A2", "A1:A2")));
  }

  @Test
  void rowAndColumnBandSelectorsValidateBounds() {
    assertEquals(
        new RowBandSelector.Span("Budget", 0, 2), new RowBandSelector.Span("Budget", 0, 2));
    assertEquals(
        new ColumnBandSelector.Span("Budget", 1, 3), new ColumnBandSelector.Span("Budget", 1, 3));
    assertEquals(
        new RowBandSelector.Insertion("Budget", 5, 2),
        new RowBandSelector.Insertion("Budget", 5, 2));
    assertEquals(
        new ColumnBandSelector.Insertion("Budget", 4, 1),
        new ColumnBandSelector.Insertion("Budget", 4, 1));

    assertThrows(IllegalArgumentException.class, () -> new RowBandSelector.Span("Budget", -1, 0));
    assertThrows(IllegalArgumentException.class, () -> new RowBandSelector.Span("Budget", 2, 1));
    assertThrows(
        IllegalArgumentException.class, () -> new ColumnBandSelector.Span("Budget", -1, 0));
    assertThrows(IllegalArgumentException.class, () -> new ColumnBandSelector.Span("Budget", 2, 1));
    assertThrows(
        IllegalArgumentException.class, () -> new RowBandSelector.Insertion("Budget", 0, 0));
    assertThrows(
        IllegalArgumentException.class, () -> new ColumnBandSelector.Insertion("Budget", 0, 0));
  }

  @Test
  void tablePivotAndNamedRangeSelectorsValidateScopeAwareIdentity() {
    assertEquals(new TableSelector.All(), new TableSelector.All());
    assertEquals(List.of("BudgetTable"), new TableSelector.ByNames(List.of("BudgetTable")).names());
    assertEquals(
        new TableSelector.ByNameOnSheet("BudgetTable", "Budget"),
        new TableSelector.ByNameOnSheet("BudgetTable", "Budget"));
    assertThrows(IllegalArgumentException.class, () -> new TableSelector.ByNames(List.of()));

    PivotTableSelector.ByNames pivots =
        new PivotTableSelector.ByNames(List.of("Sales Pivot 2026", "Ops Pivot"));
    assertEquals(List.of("Sales Pivot 2026", "Ops Pivot"), pivots.names());
    assertThrows(
        IllegalArgumentException.class,
        () -> new PivotTableSelector.ByNames(List.of("Sales Pivot 2026", "sales pivot 2026")));

    NamedRangeSelector.AnyOf selector =
        new NamedRangeSelector.AnyOf(
            List.of(
                new NamedRangeSelector.WorkbookScope("BudgetTotal"),
                new NamedRangeSelector.SheetScope("LocalItem", "Budget")));
    assertEquals(2, selector.selectors().size());
    assertInstanceOf(NamedRangeSelector.SheetScope.class, selector.selectors().get(1));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new NamedRangeSelector.AnyOf(
                List.of(
                    new NamedRangeSelector.WorkbookScope("BudgetTotal"),
                    new NamedRangeSelector.WorkbookScope("budgettotal"))));
  }

  @Test
  void tableRowAndCellSelectorsValidateComposition() {
    TableSelector.ByName table = new TableSelector.ByName("BudgetTable");
    TableRowSelector.ByKeyCell row =
        new TableRowSelector.ByKeyCell(table, "Item", new CellInput.Text(text("Laptop")));
    TableCellSelector.ByColumnName cell = new TableCellSelector.ByColumnName(row, "Amount");

    assertEquals(table, row.table());
    assertEquals("Item", row.columnName());
    assertEquals("Amount", cell.columnName());
    assertInstanceOf(
        TableRowSelector.ByKeyCell.class,
        new TableRowSelector.ByKeyCell(table, "Item", new CellInput.Blank()));
    assertInstanceOf(
        TableRowSelector.ByKeyCell.class,
        new TableRowSelector.ByKeyCell(table, "Item", new CellInput.Numeric(42.0d)));
    assertInstanceOf(
        TableRowSelector.ByKeyCell.class,
        new TableRowSelector.ByKeyCell(table, "Item", new CellInput.BooleanValue(true)));
    assertInstanceOf(
        TableRowSelector.ByKeyCell.class,
        new TableRowSelector.ByKeyCell(
            table, "Item", new CellInput.Formula(TextSourceInput.inline("A1"))));
    assertThrows(
        NullPointerException.class, () -> new TableRowSelector.ByKeyCell(table, "Item", null));
    assertThrows(
        NullPointerException.class, () -> new TableCellSelector.ByColumnName(null, "Amount"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new TableRowSelector.ByIndex(new TableSelector.All(), 0));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new TableRowSelector.ByKeyCell(
                new TableSelector.ByNames(List.of("BudgetTable")), "Item", new CellInput.Blank()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new TableRowSelector.ByKeyCell(
                table,
                "Item",
                new CellInput.RichText(List.of(new RichTextRunInput(text("Ada"), null)))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new TableRowSelector.ByKeyCell(
                table, "Item", new CellInput.Date(LocalDate.of(2026, 4, 18))));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new TableRowSelector.ByKeyCell(
                table, "Item", new CellInput.DateTime(LocalDateTime.of(2026, 4, 18, 12, 30))));
    assertThrows(
        IllegalArgumentException.class,
        () -> new TableCellSelector.ByColumnName(new TableRowSelector.AllRows(table), "Amount"));
  }

  private static TextSourceInput text(String value) {
    return TextSourceInput.inline(value);
  }
}
