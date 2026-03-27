package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import org.junit.jupiter.api.Test;

/** Tests for NamedRangeNotFoundException accessors and messages. */
class NamedRangeNotFoundExceptionTest {
  @Test
  void describesMissingWorkbookAndSheetScopedNames() {
    NamedRangeNotFoundException workbookScope =
        new NamedRangeNotFoundException("BudgetTotal", new ExcelNamedRangeScope.WorkbookScope());
    NamedRangeNotFoundException sheetScope =
        new NamedRangeNotFoundException("LocalItem", new ExcelNamedRangeScope.SheetScope("Budget"));

    assertEquals("BudgetTotal", workbookScope.name());
    assertEquals(new ExcelNamedRangeScope.WorkbookScope(), workbookScope.scope());
    assertTrue(workbookScope.getMessage().contains("workbook scope"));
    assertEquals("LocalItem", sheetScope.name());
    assertEquals(new ExcelNamedRangeScope.SheetScope("Budget"), sheetScope.scope());
    assertTrue(sheetScope.getMessage().contains("sheet scope: Budget"));
  }

  @Test
  void validatesMissingNamedRangeExceptionInputs() {
    assertThrows(
        NullPointerException.class, () -> new NamedRangeNotFoundException("BudgetTotal", null));
  }
}
