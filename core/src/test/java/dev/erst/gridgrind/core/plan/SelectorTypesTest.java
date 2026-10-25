package dev.erst.gridgrind.core.plan;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.model.style.RichTextRunInput;
import dev.erst.gridgrind.core.model.value.CellInput;
import dev.erst.gridgrind.core.source.TextSourceInput;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Validation coverage for selector-first protocol targeting types. */
class SelectorTypesTest {
  @Test
  void allSelectorFamiliesExposeExpectedCardinalityContracts() {
    assertEquals(SelectorCardinality.EXACTLY_ONE, new WorkbookSelector.Current().cardinality());
    assertEquals(SelectorCardinality.ANY_NUMBER, new SheetSelector.All().cardinality());
    assertEquals(
        SelectorCardinality.EXACTLY_ONE,
        new SheetSelector.ByName(new SheetName("Budget")).cardinality());
    assertEquals(
        SelectorCardinality.ONE_OR_MORE,
        new SheetSelector.ByNames(List.of(new SheetName("Budget"), new SheetName("Ops")))
            .cardinality());

    assertEquals(
        SelectorCardinality.ANY_NUMBER,
        new CellSelector.AllUsedInSheet(new SheetName("Budget")).cardinality());
    assertEquals(
        SelectorCardinality.EXACTLY_ONE,
        new CellSelector.ByAddress(new SheetName("Budget"), "A1").cardinality());
    assertEquals(
        SelectorCardinality.ONE_OR_MORE,
        new CellSelector.ByAddresses(new SheetName("Budget"), List.of("A1", "B2")).cardinality());

    assertEquals(
        SelectorCardinality.ANY_NUMBER,
        new RangeSelector.AllOnSheet(new SheetName("Budget")).cardinality());
    assertEquals(
        SelectorCardinality.EXACTLY_ONE,
        new RangeSelector.ByRange(new SheetName("Budget"), "A1:B2").cardinality());
    assertEquals(
        SelectorCardinality.ONE_OR_MORE,
        new RangeSelector.ByRanges(new SheetName("Budget"), List.of("A1:A2", "C3:D4"))
            .cardinality());
    assertEquals(
        SelectorCardinality.EXACTLY_ONE,
        new RangeSelector.RectangularWindow(new SheetName("Budget"), "B3", 2, 3).cardinality());

    assertEquals(
        SelectorCardinality.EXACTLY_ONE,
        new RowBandSelector.Span(new SheetName("Budget"), 0, 2).cardinality());
    assertEquals(
        SelectorCardinality.EXACTLY_ONE,
        new RowBandSelector.Insertion(new SheetName("Budget"), 5, 2).cardinality());
    assertEquals(
        SelectorCardinality.EXACTLY_ONE,
        new ColumnBandSelector.Span(new SheetName("Budget"), 1, 3).cardinality());
    assertEquals(
        SelectorCardinality.EXACTLY_ONE,
        new ColumnBandSelector.Insertion(new SheetName("Budget"), 4, 1).cardinality());

    assertEquals(SelectorCardinality.ANY_NUMBER, new TableSelector.All().cardinality());
    assertEquals(
        SelectorCardinality.EXACTLY_ONE, new TableSelector.ByName("BudgetTable").cardinality());
    assertEquals(
        SelectorCardinality.ONE_OR_MORE,
        new TableSelector.ByNames(List.of("BudgetTable", "OpsTable")).cardinality());
    assertEquals(
        SelectorCardinality.EXACTLY_ONE,
        new TableSelector.ByNameOnSheet("BudgetTable", new SheetName("Budget")).cardinality());

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
        new NamedRangeSelector.SheetScope("LocalItem", new SheetName("Budget")).cardinality());
    assertEquals(
        SelectorCardinality.ONE_OR_MORE,
        new NamedRangeSelector.AnyOf(
                List.of(
                    new NamedRangeSelector.WorkbookScope("BudgetTotal"),
                    new NamedRangeSelector.SheetScope("LocalItem", new SheetName("Budget"))))
            .cardinality());

    assertEquals(
        SelectorCardinality.ANY_NUMBER,
        new DrawingObjectSelector.AllOnSheet(new SheetName("Budget")).cardinality());
    assertEquals(
        SelectorCardinality.EXACTLY_ONE,
        new DrawingObjectSelector.ByName(new SheetName("Budget"), "OpsPicture").cardinality());
    assertEquals(
        SelectorCardinality.ANY_NUMBER,
        new ChartSelector.AllOnSheet(new SheetName("Budget")).cardinality());
    assertEquals(
        SelectorCardinality.EXACTLY_ONE,
        new ChartSelector.ByName(new SheetName("Budget"), "OpsChart").cardinality());
    assertEquals(SelectorCardinality.ANY_NUMBER, new PivotTableSelector.All().cardinality());
    assertEquals(
        SelectorCardinality.EXACTLY_ONE,
        new PivotTableSelector.ByName("Sales Pivot 2026").cardinality());
    assertEquals(
        SelectorCardinality.ONE_OR_MORE,
        new PivotTableSelector.ByNames(List.of("Sales Pivot 2026", "Ops Pivot")).cardinality());
    assertEquals(
        SelectorCardinality.EXACTLY_ONE,
        new PivotTableSelector.ByNameOnSheet("Sales Pivot 2026", new SheetName("Report"))
            .cardinality());
  }

  @Test
  void sheetCellAndRangeSelectorsValidateAndNormalize() {
    assertEquals(new SheetSelector.All(), new SheetSelector.All());
    assertEquals(
        new SheetSelector.ByName(new SheetName("Budget")),
        new SheetSelector.ByName(new SheetName("Budget")));
    assertEquals(
        List.of("Budget", "Ops"),
        new SheetSelector.ByNames(List.of(new SheetName("Budget"), new SheetName("Ops")))
            .names().stream().map(dev.erst.gridgrind.core.vocabulary.WireString::value).toList());
    assertThrows(IllegalArgumentException.class, () -> new SheetSelector.ByNames(List.of()));

    CellSelector.ByAddresses cells =
        new CellSelector.ByAddresses(new SheetName("Budget"), List.of("A1", "$B$2"));
    assertEquals("Budget", cells.sheetName().value());
    assertEquals(List.of("A1", "$B$2"), cells.addresses());
    assertThrows(
        IllegalArgumentException.class,
        () -> new CellSelector.ByAddresses(new SheetName("Budget"), List.of("A1", "a1")));

    RangeSelector.RectangularWindow window =
        new RangeSelector.RectangularWindow(new SheetName("Budget"), "B3", 2, 3);
    assertEquals("B3:D4", window.range());
    assertThrows(
        IllegalArgumentException.class,
        () -> new RangeSelector.ByRanges(new SheetName("Budget"), List.of("A1:A2", "A1:A2")));
  }

  @Test
  void rowAndColumnBandSelectorsValidateBounds() {
    assertEquals(
        new RowBandSelector.Span(new SheetName("Budget"), 0, 2),
        new RowBandSelector.Span(new SheetName("Budget"), 0, 2));
    assertEquals(
        new ColumnBandSelector.Span(new SheetName("Budget"), 1, 3),
        new ColumnBandSelector.Span(new SheetName("Budget"), 1, 3));
    assertEquals(
        new RowBandSelector.Insertion(new SheetName("Budget"), 5, 2),
        new RowBandSelector.Insertion(new SheetName("Budget"), 5, 2));
    assertEquals(
        new ColumnBandSelector.Insertion(new SheetName("Budget"), 4, 1),
        new ColumnBandSelector.Insertion(new SheetName("Budget"), 4, 1));

    assertThrows(
        IllegalArgumentException.class,
        () -> new RowBandSelector.Span(new SheetName("Budget"), -1, 0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new RowBandSelector.Span(new SheetName("Budget"), 2, 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ColumnBandSelector.Span(new SheetName("Budget"), -1, 0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ColumnBandSelector.Span(new SheetName("Budget"), 2, 1));
    assertThrows(
        IllegalArgumentException.class,
        () -> new RowBandSelector.Insertion(new SheetName("Budget"), 0, 0));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ColumnBandSelector.Insertion(new SheetName("Budget"), 0, 0));
  }

  @Test
  void tablePivotAndNamedRangeSelectorsValidateScopeAwareIdentity() {
    assertEquals(new TableSelector.All(), new TableSelector.All());
    assertEquals(List.of("BudgetTable"), new TableSelector.ByNames(List.of("BudgetTable")).names());
    assertEquals(
        new TableSelector.ByNameOnSheet("BudgetTable", new SheetName("Budget")),
        new TableSelector.ByNameOnSheet("BudgetTable", new SheetName("Budget")));
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
                new NamedRangeSelector.SheetScope("LocalItem", new SheetName("Budget"))));
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
        new TableRowSelector.ByKeyCell(table, "Item", new CellInput.NumberValue(42.0d)));
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
                new CellInput.RichText(
                    List.of(new RichTextRunInput(text("Ada"), Optional.empty())))));
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
