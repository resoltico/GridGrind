package dev.erst.gridgrind.engine.runtime;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.model.formula.NamedRangeScopeReport;
import dev.erst.gridgrind.core.plan.NamedRangeSelector;
import dev.erst.gridgrind.core.vocabulary.DefinedName;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import org.junit.jupiter.api.Test;

/** Tests for NamedRangeScope record construction and engine conversion. */
class NamedRangeScopeTest {
  @Test
  void convertsSelectorScopedNamedRangeDeletionTargets() {
    assertEquals(
        new NamedRangeScopeReport.Workbook(),
        WorkbookCommandLayoutInputConverter.toExcelNamedRangeScope(
            new NamedRangeSelector.WorkbookScope(new DefinedName("BudgetTotal"))));
    assertEquals(
        new NamedRangeScopeReport.Sheet(
            new dev.erst.gridgrind.core.vocabulary.SheetLabel("Budget")),
        WorkbookCommandLayoutInputConverter.toExcelNamedRangeScope(
            new NamedRangeSelector.SheetScope(
                new DefinedName("LocalItem"), new SheetName("Budget"))));
    assertEquals(
        "BudgetTotal",
        WorkbookCommandLayoutInputConverter.toExcelNamedRangeName(
            new NamedRangeSelector.WorkbookScope(new DefinedName("BudgetTotal"))));
    assertEquals(
        "LocalItem",
        WorkbookCommandLayoutInputConverter.toExcelNamedRangeName(
            new NamedRangeSelector.SheetScope(
                new DefinedName("LocalItem"), new SheetName("Budget"))));
  }
}
