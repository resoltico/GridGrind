package dev.erst.gridgrind.core.model.formula;

import static org.junit.jupiter.api.Assertions.*;

import org.junit.jupiter.api.Test;

/** Tests for NamedRangeReport record validation and typed accessors. */
class NamedRangeReportTest {
  @Test
  void constructsRangeAndFormulaNamedRangeReports() {
    NamedRangeReport.RangeReport rangeSnapshot =
        new NamedRangeReport.RangeReport(
            new dev.erst.gridgrind.core.vocabulary.DefinedNameLabel("BudgetTotal"),
            new NamedRangeScopeReport.Workbook(),
            "Budget!$B$4",
            NamedRangeTarget.range(
                new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
                new dev.erst.gridgrind.core.vocabulary.RangeRef("B4")));
    NamedRangeReport.FormulaReport formulaSnapshot =
        new NamedRangeReport.FormulaReport(
            new dev.erst.gridgrind.core.vocabulary.DefinedNameLabel("BudgetRollup"),
            new NamedRangeScopeReport.Sheet(
                new dev.erst.gridgrind.core.vocabulary.SheetLabel("Budget")),
            "SUM(Budget!$B$2:$B$3)");

    assertEquals("BudgetTotal", rangeSnapshot.name().value());
    assertEquals("BudgetRollup", formulaSnapshot.name().value());
    assertEquals(
        new NamedRangeScopeReport.Sheet(
            new dev.erst.gridgrind.core.vocabulary.SheetLabel("Budget")),
        formulaSnapshot.scope());
    assertEquals("SUM(Budget!$B$2:$B$3)", formulaSnapshot.refersToFormula());
  }

  @Test
  void validatesRangeAndFormulaNamedRangeReports() {
    assertThrows(
        NullPointerException.class,
        () ->
            new NamedRangeReport.RangeReport(
                new dev.erst.gridgrind.core.vocabulary.DefinedNameLabel("BudgetTotal"),
                null,
                "Budget!$B$4",
                NamedRangeTarget.range(
                    new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
                    new dev.erst.gridgrind.core.vocabulary.RangeRef("B4"))));
    assertThrows(
        NullPointerException.class,
        () ->
            new NamedRangeReport.RangeReport(
                new dev.erst.gridgrind.core.vocabulary.DefinedNameLabel("BudgetTotal"),
                new NamedRangeScopeReport.Workbook(),
                null,
                NamedRangeTarget.range(
                    new dev.erst.gridgrind.core.vocabulary.SheetName("Budget"),
                    new dev.erst.gridgrind.core.vocabulary.RangeRef("B4"))));
    assertThrows(
        NullPointerException.class,
        () ->
            new NamedRangeReport.RangeReport(
                new dev.erst.gridgrind.core.vocabulary.DefinedNameLabel("BudgetTotal"),
                new NamedRangeScopeReport.Workbook(),
                "Budget!$B$4",
                null));
    assertThrows(
        NullPointerException.class,
        () ->
            new NamedRangeReport.FormulaReport(
                new dev.erst.gridgrind.core.vocabulary.DefinedNameLabel("BudgetRollup"),
                null,
                "SUM(Budget!$B$2:$B$3)"));
    assertThrows(
        NullPointerException.class,
        () ->
            new NamedRangeReport.FormulaReport(
                new dev.erst.gridgrind.core.vocabulary.DefinedNameLabel("BudgetRollup"),
                new NamedRangeScopeReport.Workbook(),
                null));
  }
}
