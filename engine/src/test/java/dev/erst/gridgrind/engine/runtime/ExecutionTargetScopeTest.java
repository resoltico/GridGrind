package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.plan.CellSelector;
import dev.erst.gridgrind.core.plan.ChartSelector;
import dev.erst.gridgrind.core.plan.ColumnBandSelector;
import dev.erst.gridgrind.core.plan.DrawingObjectSelector;
import dev.erst.gridgrind.core.plan.NamedRangeSelector;
import dev.erst.gridgrind.core.plan.PivotTableSelector;
import dev.erst.gridgrind.core.plan.RangeSelector;
import dev.erst.gridgrind.core.plan.RowBandSelector;
import dev.erst.gridgrind.core.plan.Selector;
import dev.erst.gridgrind.core.plan.SheetSelector;
import dev.erst.gridgrind.core.plan.TableCellSelector;
import dev.erst.gridgrind.core.plan.TableRowSelector;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.core.plan.WorkbookSelector;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.DefinedName;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.TableName;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Verifies exact selected-sheet extraction across every selector variant. */
class ExecutionTargetScopeTest {
  @Test
  void derivesExactSheetNamesOnlyForSelectorsThatStateThemCompletely() {
    List<Selector> budgetSelectors =
        List.of(
            new SheetSelector.ByName(new SheetName("Budget")),
            new CellSelector.AllUsedInSheet(new SheetName("Budget")),
            new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
            new CellSelector.ByAddresses(new SheetName("Budget"), List.of(new CellAddress("A1"))),
            new RangeSelector.AllOnSheet(new SheetName("Budget")),
            new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("A1:B2")),
            new RangeSelector.ByRanges(
                new SheetName("Budget"), List.of(new RangeRef("A1:A1"), new RangeRef("B1:B1"))),
            new RangeSelector.RectangularWindow(
                new SheetName("Budget"), new CellAddress("A1"), 1, 1),
            new RowBandSelector.Span(new SheetName("Budget"), 0, 0),
            new RowBandSelector.Insertion(new SheetName("Budget"), 0, 1),
            new ColumnBandSelector.Span(new SheetName("Budget"), 0, 0),
            new ColumnBandSelector.Insertion(new SheetName("Budget"), 0, 1),
            new DrawingObjectSelector.AllOnSheet(new SheetName("Budget")),
            new DrawingObjectSelector.ByName(new SheetName("Budget"), "Logo"),
            new ChartSelector.AllOnSheet(new SheetName("Budget")),
            new ChartSelector.ByName(new SheetName("Budget"), "Revenue"),
            new TableSelector.ByNameOnSheet(new TableName("Ledger"), new SheetName("Budget")),
            new PivotTableSelector.ByNameOnSheet("Pivot", new SheetName("Budget")),
            new NamedRangeSelector.SheetScope(new DefinedName("Total"), new SheetName("Budget")),
            new TableRowSelector.AllRows(
                new TableSelector.ByNameOnSheet(new TableName("Ledger"), new SheetName("Budget"))),
            new TableRowSelector.ByIndex(
                new TableSelector.ByNameOnSheet(new TableName("Ledger"), new SheetName("Budget")),
                0),
            new TableRowSelector.ByKeyCell(
                new TableSelector.ByNameOnSheet(new TableName("Ledger"), new SheetName("Budget")),
                "Owner",
                new CellValueSpec.Blank()),
            new TableCellSelector.ByColumnName(
                new TableRowSelector.ByIndex(
                    new TableSelector.ByNameOnSheet(
                        new TableName("Ledger"), new SheetName("Budget")),
                    0),
                "Amount"));

    for (Selector selector : budgetSelectors) {
      assertEquals(Optional.of(List.of("Budget")), ExecutionTargetScope.exactSheetNames(selector));
    }
    assertEquals(
        Optional.of(List.of("Budget", "Ledger")),
        ExecutionTargetScope.exactSheetNames(
            new SheetSelector.ByNames(List.of(new SheetName("Budget"), new SheetName("Ledger")))));
    for (Selector selector :
        List.of(
            new SheetSelector.All(),
            new TableSelector.All(),
            new TableSelector.ByName(new TableName("Ledger")),
            new TableSelector.ByNames(List.of(new TableName("Ledger"))),
            new PivotTableSelector.All(),
            new PivotTableSelector.ByName("Pivot"),
            new PivotTableSelector.ByNames(List.of("Pivot")),
            new NamedRangeSelector.All(),
            new NamedRangeSelector.AnyOf(
                List.of(new NamedRangeSelector.ByName(new DefinedName("Total")))),
            new NamedRangeSelector.ByName(new DefinedName("Total")),
            new NamedRangeSelector.ByNames(List.of(new DefinedName("Total"))),
            new NamedRangeSelector.WorkbookScope(new DefinedName("Total")),
            new WorkbookSelector.Current())) {
      assertEquals(Optional.empty(), ExecutionTargetScope.exactSheetNames(selector));
    }
  }
}
