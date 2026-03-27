package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import org.junit.jupiter.api.Test;

/** Tests for ExcelNamedRangeSelector constructor invariants. */
class ExcelNamedRangeSelectorTest {
  @Test
  void createsEachSelectorVariant() {
    ExcelNamedRangeSelector.ByName byName = new ExcelNamedRangeSelector.ByName("BudgetTotal");
    ExcelNamedRangeSelector.WorkbookScope workbookScope =
        new ExcelNamedRangeSelector.WorkbookScope("BudgetTotal");
    ExcelNamedRangeSelector.SheetScope sheetScope =
        new ExcelNamedRangeSelector.SheetScope("LocalItem", "Budget");

    assertEquals("BudgetTotal", byName.name());
    assertEquals("BudgetTotal", workbookScope.name());
    assertEquals("LocalItem", sheetScope.name());
    assertEquals("Budget", sheetScope.sheetName());
  }

  @Test
  void rejectsBlankSelectorFields() {
    assertThrows(NullPointerException.class, () -> new ExcelNamedRangeSelector.ByName(null));
    assertThrows(IllegalArgumentException.class, () -> new ExcelNamedRangeSelector.ByName(" "));
    assertThrows(NullPointerException.class, () -> new ExcelNamedRangeSelector.WorkbookScope(null));
    assertThrows(
        IllegalArgumentException.class, () -> new ExcelNamedRangeSelector.WorkbookScope(" "));
    assertThrows(
        NullPointerException.class,
        () -> new ExcelNamedRangeSelector.SheetScope("LocalItem", null));
    assertThrows(
        IllegalArgumentException.class,
        () -> new ExcelNamedRangeSelector.SheetScope("LocalItem", " "));
  }
}
