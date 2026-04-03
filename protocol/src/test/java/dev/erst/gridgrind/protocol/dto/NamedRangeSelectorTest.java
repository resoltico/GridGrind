package dev.erst.gridgrind.protocol.dto;

import static org.junit.jupiter.api.Assertions.*;

import org.junit.jupiter.api.Test;

/** Tests for NamedRangeSelector record construction. */
class NamedRangeSelectorTest {
  @Test
  void buildsSupportedNamedRangeSelectors() {
    assertEquals("BudgetTotal", new NamedRangeSelector.ByName("BudgetTotal").name());
    assertEquals("BudgetTotal", new NamedRangeSelector.WorkbookScope("BudgetTotal").name());
    assertEquals("Budget", new NamedRangeSelector.SheetScope("BudgetTotal", "Budget").sheetName());
  }

  @Test
  void validatesSelectorInputs() {
    assertThrows(NullPointerException.class, () -> new NamedRangeSelector.ByName(null));
    assertThrows(IllegalArgumentException.class, () -> new NamedRangeSelector.ByName(" "));
    assertThrows(NullPointerException.class, () -> new NamedRangeSelector.WorkbookScope(null));
    assertThrows(
        NullPointerException.class, () -> new NamedRangeSelector.SheetScope("BudgetTotal", null));
    assertThrows(
        IllegalArgumentException.class, () -> new NamedRangeSelector.SheetScope(" ", "Budget"));
    assertThrows(
        IllegalArgumentException.class,
        () -> new NamedRangeSelector.SheetScope("BudgetTotal", " "));
  }
}
