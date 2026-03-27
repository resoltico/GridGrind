package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import java.util.ArrayList;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for ExcelNamedRangeSelection constructor invariants and defensive copies. */
class ExcelNamedRangeSelectionTest {
  @Test
  void createsAllSelection() {
    ExcelNamedRangeSelection.All selection = new ExcelNamedRangeSelection.All();

    assertInstanceOf(ExcelNamedRangeSelection.All.class, selection);
  }

  @Test
  void selectedCopiesSelectorsAndPreservesOrder() {
    List<ExcelNamedRangeSelector> selectors =
        new ArrayList<>(
            List.of(
                new ExcelNamedRangeSelector.ByName("BudgetTotal"),
                new ExcelNamedRangeSelector.SheetScope("LocalItem", "Budget")));

    ExcelNamedRangeSelection.Selected selection = new ExcelNamedRangeSelection.Selected(selectors);
    selectors.clear();

    assertEquals(
        List.of(
            new ExcelNamedRangeSelector.ByName("BudgetTotal"),
            new ExcelNamedRangeSelector.SheetScope("LocalItem", "Budget")),
        selection.selectors());
  }

  @Test
  void selectedRejectsInvalidSelectorCollections() {
    assertThrows(NullPointerException.class, () -> new ExcelNamedRangeSelection.Selected(null));
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelNamedRangeSelection.Selected(List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new ExcelNamedRangeSelection.Selected(
                List.of(
                    new ExcelNamedRangeSelector.ByName("BudgetTotal"),
                    new ExcelNamedRangeSelector.ByName("BudgetTotal"))));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelNamedRangeSelection.Selected(
                List.of(new ExcelNamedRangeSelector.ByName("BudgetTotal"), null)));
  }
}
