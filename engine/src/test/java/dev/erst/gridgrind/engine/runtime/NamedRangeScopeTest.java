package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.plan.NamedRangeSelector;
import dev.erst.gridgrind.core.vocabulary.NamedRangeScope;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.excel.ExcelNamedRangeScope;
import org.junit.jupiter.api.Test;

/** Tests for NamedRangeScope record construction and engine conversion. */
class NamedRangeScopeTest {
  @Test
  void convertsWorkbookAndSheetScopes() {
    assertEquals(
        new ExcelNamedRangeScope.WorkbookScope(),
        WorkbookCommandLayoutInputConverter.toExcelNamedRangeScope(new NamedRangeScope.Workbook()));
    assertEquals(
        new ExcelNamedRangeScope.SheetScope("Budget"),
        WorkbookCommandLayoutInputConverter.toExcelNamedRangeScope(
            new NamedRangeScope.Sheet(new SheetName("Budget"))));
  }

  @Test
  void convertsSelectorScopedNamedRangeDeletionTargets() {
    assertEquals(
        new ExcelNamedRangeScope.WorkbookScope(),
        WorkbookCommandLayoutInputConverter.toExcelNamedRangeScope(
            new NamedRangeSelector.WorkbookScope("BudgetTotal")));
    assertEquals(
        new ExcelNamedRangeScope.SheetScope("Budget"),
        WorkbookCommandLayoutInputConverter.toExcelNamedRangeScope(
            new NamedRangeSelector.SheetScope("LocalItem", new SheetName("Budget"))));
    assertEquals(
        "BudgetTotal",
        WorkbookCommandLayoutInputConverter.toExcelNamedRangeName(
            new NamedRangeSelector.WorkbookScope("BudgetTotal")));
    assertEquals(
        "LocalItem",
        WorkbookCommandLayoutInputConverter.toExcelNamedRangeName(
            new NamedRangeSelector.SheetScope("LocalItem", new SheetName("Budget"))));
  }

  @Test
  void validatesSheetScopeInput() {
    assertThrows(NullPointerException.class, () -> new NamedRangeScope.Sheet(null));
    assertThrows(
        IllegalArgumentException.class, () -> new NamedRangeScope.Sheet(new SheetName(" ")));
  }
}
