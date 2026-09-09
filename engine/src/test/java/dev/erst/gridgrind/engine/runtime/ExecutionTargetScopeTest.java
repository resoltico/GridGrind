package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.ChartSelector;
import dev.erst.gridgrind.contract.selector.ColumnBandSelector;
import dev.erst.gridgrind.contract.selector.DrawingObjectSelector;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.contract.selector.PivotTableSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.RowBandSelector;
import dev.erst.gridgrind.contract.selector.Selector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.TableCellSelector;
import dev.erst.gridgrind.contract.selector.TableRowSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Verifies exact selected-sheet extraction across every selector variant. */
class ExecutionTargetScopeTest {
  @Test
  void derivesExactSheetNamesOnlyForSelectorsThatStateThemCompletely() {
    List<Selector> budgetSelectors =
        List.of(
            new SheetSelector.ByName("Budget"),
            new CellSelector.AllUsedInSheet("Budget"),
            new CellSelector.ByAddress("Budget", "A1"),
            new CellSelector.ByAddresses("Budget", List.of("A1")),
            new RangeSelector.AllOnSheet("Budget"),
            new RangeSelector.ByRange("Budget", "A1:B2"),
            new RangeSelector.ByRanges("Budget", List.of("A1:A1", "B1:B1")),
            new RangeSelector.RectangularWindow("Budget", "A1", 1, 1),
            new RowBandSelector.Span("Budget", 0, 0),
            new RowBandSelector.Insertion("Budget", 0, 1),
            new ColumnBandSelector.Span("Budget", 0, 0),
            new ColumnBandSelector.Insertion("Budget", 0, 1),
            new DrawingObjectSelector.AllOnSheet("Budget"),
            new DrawingObjectSelector.ByName("Budget", "Logo"),
            new ChartSelector.AllOnSheet("Budget"),
            new ChartSelector.ByName("Budget", "Revenue"),
            new TableSelector.ByNameOnSheet("Ledger", "Budget"),
            new PivotTableSelector.ByNameOnSheet("Pivot", "Budget"),
            new NamedRangeSelector.SheetScope("Total", "Budget"),
            new TableRowSelector.AllRows(new TableSelector.ByNameOnSheet("Ledger", "Budget")),
            new TableRowSelector.ByIndex(new TableSelector.ByNameOnSheet("Ledger", "Budget"), 0),
            new TableRowSelector.ByKeyCell(
                new TableSelector.ByNameOnSheet("Ledger", "Budget"),
                "Owner",
                new CellInput.Blank()),
            new TableCellSelector.ByColumnName(
                new TableRowSelector.ByIndex(
                    new TableSelector.ByNameOnSheet("Ledger", "Budget"), 0),
                "Amount"));

    for (Selector selector : budgetSelectors) {
      assertEquals(Optional.of(List.of("Budget")), ExecutionTargetScope.exactSheetNames(selector));
    }
    assertEquals(
        Optional.of(List.of("Budget", "Ledger")),
        ExecutionTargetScope.exactSheetNames(
            new SheetSelector.ByNames(List.of("Budget", "Ledger"))));
    for (Selector selector :
        List.of(
            new SheetSelector.All(),
            new TableSelector.All(),
            new TableSelector.ByName("Ledger"),
            new TableSelector.ByNames(List.of("Ledger")),
            new PivotTableSelector.All(),
            new PivotTableSelector.ByName("Pivot"),
            new PivotTableSelector.ByNames(List.of("Pivot")),
            new NamedRangeSelector.All(),
            new NamedRangeSelector.AnyOf(List.of(new NamedRangeSelector.ByName("Total"))),
            new NamedRangeSelector.ByName("Total"),
            new NamedRangeSelector.ByNames(List.of("Total")),
            new NamedRangeSelector.WorkbookScope("Total"),
            new WorkbookSelector.Current())) {
      assertEquals(Optional.empty(), ExecutionTargetScope.exactSheetNames(selector));
    }
  }
}
