package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.core.vocabulary.TableName;
import dev.erst.gridgrind.excel.ExcelTableSelection;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for selector-facing table invariants and conversion. */
class TableSelectionTest {
  @Test
  void convertsAllAndByNamesSelectors() {
    assertEquals(
        new ExcelTableSelection.All(),
        SelectorConverter.toExcelTableSelection(new TableSelector.All()));
    assertEquals(
        new ExcelTableSelection.ByNames(List.of("BudgetTable")),
        SelectorConverter.toExcelTableSelection(
            new TableSelector.ByNames(List.of(new TableName("BudgetTable")))));
  }

  @Test
  void validatesByNamesSelection() {
    assertThrows(IllegalArgumentException.class, () -> new TableSelector.ByNames(List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new TableSelector.ByNames(
                List.of(new TableName("BudgetTable"), new TableName("budgettable"))));
    assertThrows(
        NullPointerException.class,
        () -> new TableSelector.ByNames(List.of(new TableName("BudgetTable"), null)));
  }
}
