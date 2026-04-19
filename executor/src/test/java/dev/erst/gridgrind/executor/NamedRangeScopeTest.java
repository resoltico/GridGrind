package dev.erst.gridgrind.executor;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.contract.dto.NamedRangeScope;
import dev.erst.gridgrind.contract.selector.NamedRangeSelector;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import org.junit.jupiter.api.Test;

/** Tests for NamedRangeScope record construction and engine conversion. */
class NamedRangeScopeTest {
  @Test
  void convertsWorkbookAndSheetScopes() {
    assertEquals(
        new ExcelNamedRangeScope.WorkbookScope(),
        WorkbookCommandConverter.toExcelNamedRangeScope(new NamedRangeScope.Workbook()));
    assertEquals(
        new ExcelNamedRangeScope.SheetScope("Budget"),
        WorkbookCommandConverter.toExcelNamedRangeScope(new NamedRangeScope.Sheet("Budget")));
  }

  @Test
  void convertsSelectorScopedNamedRangeDeletionTargets() {
    assertEquals(
        new ExcelNamedRangeScope.WorkbookScope(),
        WorkbookCommandConverter.toExcelNamedRangeScope(
            new NamedRangeSelector.WorkbookScope("BudgetTotal")));
    assertEquals(
        new ExcelNamedRangeScope.SheetScope("Budget"),
        WorkbookCommandConverter.toExcelNamedRangeScope(
            new NamedRangeSelector.SheetScope("LocalItem", "Budget")));
    assertEquals(
        "BudgetTotal",
        WorkbookCommandConverter.toExcelNamedRangeName(
            new NamedRangeSelector.WorkbookScope("BudgetTotal")));
    assertEquals(
        "LocalItem",
        WorkbookCommandConverter.toExcelNamedRangeName(
            new NamedRangeSelector.SheetScope("LocalItem", "Budget")));
  }

  @Test
  void validatesSheetScopeInput() {
    assertThrows(NullPointerException.class, () -> new NamedRangeScope.Sheet(null));
    assertThrows(IllegalArgumentException.class, () -> new NamedRangeScope.Sheet(" "));
  }
}
