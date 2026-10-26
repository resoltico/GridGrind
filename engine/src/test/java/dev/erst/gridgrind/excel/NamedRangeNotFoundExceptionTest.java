package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.formula.NamedRangeScopeReport;
import org.junit.jupiter.api.Test;

/** Tests for NamedRangeNotFoundException accessors and messages. */
class NamedRangeNotFoundExceptionTest {
  @Test
  void describesMissingWorkbookAndSheetScopedNames() {
    NamedRangeNotFoundException workbookScope =
        new NamedRangeNotFoundException("BudgetTotal", new NamedRangeScopeReport.Workbook());
    NamedRangeNotFoundException sheetScope =
        new NamedRangeNotFoundException(
            "LocalItem",
            new NamedRangeScopeReport.Sheet(
                new dev.erst.gridgrind.core.vocabulary.SheetLabel("Budget")));

    assertEquals("BudgetTotal", workbookScope.name());
    assertEquals(new NamedRangeScopeReport.Workbook(), workbookScope.scope());
    assertTrue(workbookScope.getMessage().contains("workbook scope"));
    assertEquals("LocalItem", sheetScope.name());
    assertEquals(
        new NamedRangeScopeReport.Sheet(
            new dev.erst.gridgrind.core.vocabulary.SheetLabel("Budget")),
        sheetScope.scope());
    assertTrue(sheetScope.getMessage().contains("sheet scope: Budget"));
  }

  @Test
  void validatesMissingNamedRangeExceptionInputs() {
    assertThrows(
        NullPointerException.class, () -> new NamedRangeNotFoundException("BudgetTotal", null));
  }
}
