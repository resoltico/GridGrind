package dev.erst.gridgrind.protocol;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.ExcelTableSelection;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for protocol-facing table selection invariants and conversion. */
class TableSelectionTest {
  @Test
  void convertsAllAndByNamesSelectors() {
    assertEquals(new ExcelTableSelection.All(), new TableSelection.All().toExcelTableSelection());
    assertEquals(
        new ExcelTableSelection.ByNames(List.of("BudgetTable")),
        new TableSelection.ByNames(List.of("BudgetTable")).toExcelTableSelection());
  }

  @Test
  void validatesByNamesSelection() {
    assertThrows(IllegalArgumentException.class, () -> new TableSelection.ByNames(List.of()));
    assertThrows(
        IllegalArgumentException.class,
        () -> new TableSelection.ByNames(List.of("BudgetTable", "budgettable")));
    assertThrows(
        NullPointerException.class, () -> new TableSelection.ByNames(List.of("BudgetTable", null)));
  }
}
