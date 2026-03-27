package dev.erst.gridgrind.excel;

import static org.junit.jupiter.api.Assertions.*;

import org.junit.jupiter.api.Test;

/** Tests for ExcelNamedRangeSnapshot record validation and typed accessors. */
class ExcelNamedRangeSnapshotTest {
  @Test
  void constructsRangeAndFormulaNamedRangeSnapshots() {
    ExcelNamedRangeSnapshot.RangeSnapshot rangeSnapshot =
        new ExcelNamedRangeSnapshot.RangeSnapshot(
            "BudgetTotal",
            new ExcelNamedRangeScope.WorkbookScope(),
            "Budget!$B$4",
            new ExcelNamedRangeTarget("Budget", "B4"));
    ExcelNamedRangeSnapshot.FormulaSnapshot formulaSnapshot =
        new ExcelNamedRangeSnapshot.FormulaSnapshot(
            "BudgetRollup", new ExcelNamedRangeScope.SheetScope("Budget"), "SUM(Budget!$B$2:$B$3)");

    assertEquals("BudgetTotal", rangeSnapshot.name());
    assertEquals("BudgetRollup", formulaSnapshot.name());
    assertEquals(new ExcelNamedRangeScope.SheetScope("Budget"), formulaSnapshot.scope());
    assertEquals("SUM(Budget!$B$2:$B$3)", formulaSnapshot.refersToFormula());
  }

  @Test
  void validatesRangeAndFormulaNamedRangeSnapshots() {
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelNamedRangeSnapshot.RangeSnapshot(
                "BudgetTotal", null, "Budget!$B$4", new ExcelNamedRangeTarget("Budget", "B4")));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelNamedRangeSnapshot.RangeSnapshot(
                "BudgetTotal",
                new ExcelNamedRangeScope.WorkbookScope(),
                null,
                new ExcelNamedRangeTarget("Budget", "B4")));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelNamedRangeSnapshot.RangeSnapshot(
                "BudgetTotal", new ExcelNamedRangeScope.WorkbookScope(), "Budget!$B$4", null));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelNamedRangeSnapshot.FormulaSnapshot(
                "BudgetRollup", null, "SUM(Budget!$B$2:$B$3)"));
    assertThrows(
        NullPointerException.class,
        () ->
            new ExcelNamedRangeSnapshot.FormulaSnapshot(
                "BudgetRollup", new ExcelNamedRangeScope.WorkbookScope(), null));
  }
}
