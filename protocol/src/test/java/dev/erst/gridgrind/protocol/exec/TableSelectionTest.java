package dev.erst.gridgrind.protocol.exec;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.excel.ExcelTableSelection;
import dev.erst.gridgrind.protocol.dto.TableSelection;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Tests for protocol-facing table selection invariants and conversion. */
class TableSelectionTest {
  @Test
  void convertsAllAndByNamesSelectors() {
    assertEquals(
        new ExcelTableSelection.All(),
        DefaultGridGrindRequestExecutor.toExcelTableSelection(new TableSelection.All()));
    assertEquals(
        new ExcelTableSelection.ByNames(List.of("BudgetTable")),
        DefaultGridGrindRequestExecutor.toExcelTableSelection(
            new TableSelection.ByNames(List.of("BudgetTable"))));
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
