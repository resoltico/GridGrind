package dev.erst.gridgrind.contract.dto;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;

import org.junit.jupiter.api.Test;

/** Direct coverage for location-merging helpers behind the typed problem-context model. */
class ProblemContextSupportTest {
  @Test
  void mergeLocationComposesSheetCellRangeNamedRangeAndFormulaVariants() {
    assertEquals(
        new ProblemContext.ProblemLocation.Cell("Budget", "B4"),
        ProblemContext.mergeLocation(
            ProblemContext.ProblemLocation.sheet("Budget"),
            ProblemContext.ProblemLocation.address("B4")));
    assertEquals(
        new ProblemContext.ProblemLocation.Range("Budget", "A1:B2"),
        ProblemContext.mergeLocation(
            ProblemContext.ProblemLocation.sheet("Budget"),
            ProblemContext.ProblemLocation.range("A1:B2")));
    assertEquals(
        new ProblemContext.ProblemLocation.FormulaCell("Budget", "B4", "SUM(B2:B3)"),
        ProblemContext.mergeLocation(
            ProblemContext.ProblemLocation.sheet("Budget"),
            ProblemContext.ProblemLocation.formulaCell("Ignored", "B4", "SUM(B2:B3)")));
    assertEquals(
        new ProblemContext.ProblemLocation.FormulaCell("Budget", "B4", "SUM(B2:B3)"),
        ProblemContext.mergeLocation(
            ProblemContext.ProblemLocation.cell("Budget", "B4"),
            ProblemContext.ProblemLocation.formulaCell("Archive", "C9", "SUM(B2:B3)")));
    assertEquals(
        new ProblemContext.ProblemLocation.Cell("Budget", "B4"),
        ProblemContext.mergeLocation(
            ProblemContext.ProblemLocation.address("B4"),
            ProblemContext.ProblemLocation.sheet("Budget")));
    assertEquals(
        new ProblemContext.ProblemLocation.FormulaCell("Budget", "B4", "SUM(B2:B3)"),
        ProblemContext.mergeLocation(
            ProblemContext.ProblemLocation.address("B4"),
            ProblemContext.ProblemLocation.formulaCell("Budget", "C5", "SUM(B2:B3)")));
    assertEquals(
        new ProblemContext.ProblemLocation.Range("Budget", "A1:B2"),
        ProblemContext.mergeLocation(
            ProblemContext.ProblemLocation.range("A1:B2"),
            ProblemContext.ProblemLocation.sheet("Budget")));
    assertEquals(
        new ProblemContext.ProblemLocation.SheetNamedRange("Budget", "BudgetTotal"),
        ProblemContext.mergeLocation(
            ProblemContext.ProblemLocation.range("A1:B2"),
            ProblemContext.ProblemLocation.namedRange("Budget", "BudgetTotal")));
    assertEquals(
        ProblemContext.ProblemLocation.address("B4"),
        ProblemContext.mergeLocation(
            ProblemContext.ProblemLocation.unknown(),
            ProblemContext.ProblemLocation.address("B4")));
    assertEquals(
        ProblemContext.ProblemLocation.sheet("Budget"),
        ProblemContext.mergeLocation(
            ProblemContext.ProblemLocation.sheet("Budget"),
            ProblemContext.ProblemLocation.sheet("Archive")));
    assertEquals(
        new ProblemContext.ProblemLocation.Cell("Budget", "C7"),
        ProblemContext.mergeLocation(
            ProblemContext.ProblemLocation.sheet("Budget"),
            ProblemContext.ProblemLocation.cell("Archive", "C7")));
    assertEquals(
        ProblemContext.ProblemLocation.namedRange("BudgetTotal"),
        ProblemContext.mergeLocation(
            ProblemContext.ProblemLocation.sheet("Budget"),
            ProblemContext.ProblemLocation.namedRange("BudgetTotal")));
    assertEquals(
        new ProblemContext.ProblemLocation.Range("Budget", "C1:D4"),
        ProblemContext.mergeLocation(
            ProblemContext.ProblemLocation.sheet("Budget"),
            new ProblemContext.ProblemLocation.Range("Archive", "C1:D4")));
    assertEquals(
        new ProblemContext.ProblemLocation.SheetNamedRange("Ops", "BudgetTotal"),
        ProblemContext.mergeLocation(
            ProblemContext.ProblemLocation.sheet("Budget"),
            new ProblemContext.ProblemLocation.SheetNamedRange("Ops", "BudgetTotal")));
  }

  @Test
  void mergeLocationPreservesSpecificCurrentLocationsAndRejectsBlankValues() {
    ProblemContext.ProblemLocation current =
        ProblemContext.ProblemLocation.formulaCell("Budget", "B4", "SUM(B2:B3)");

    assertEquals(
        current,
        ProblemContext.mergeLocation(current, ProblemContext.ProblemLocation.sheet("Archive")));
    assertInstanceOf(
        ProblemContext.ProblemLocation.NamedRange.class,
        ProblemContext.mergeLocation(
            ProblemContext.ProblemLocation.sheet("Budget"),
            ProblemContext.ProblemLocation.namedRange("BudgetTotal")));
    assertEquals(
        ProblemContext.ProblemLocation.namedRange("Budget", "BudgetTotal"),
        ProblemContext.mergeLocation(
            ProblemContext.ProblemLocation.namedRange("Budget", "BudgetTotal"),
            ProblemContext.ProblemLocation.cell("Budget", "B4")));
    assertEquals(
        ProblemContext.ProblemLocation.namedRange("BudgetTotal"),
        ProblemContext.mergeLocation(
            ProblemContext.ProblemLocation.namedRange("BudgetTotal"),
            ProblemContext.ProblemLocation.cell("Budget", "B4")));
    assertEquals(
        new ProblemContext.ProblemLocation.Range("Budget", "A1:B2"),
        ProblemContext.mergeLocation(
            new ProblemContext.ProblemLocation.Range("Budget", "A1:B2"),
            ProblemContext.ProblemLocation.namedRange("BudgetTotal")));
    assertEquals(
        new ProblemContext.ProblemLocation.Cell("Budget", "B4"),
        ProblemContext.mergeLocation(
            ProblemContext.ProblemLocation.cell("Budget", "B4"),
            ProblemContext.ProblemLocation.sheet("Archive")));
    assertEquals(
        ProblemContext.ProblemLocation.address("B4"),
        ProblemContext.mergeLocation(
            ProblemContext.ProblemLocation.address("B4"),
            ProblemContext.ProblemLocation.range("A1:B2")));
    assertEquals(
        ProblemContext.ProblemLocation.range("A1:B2"),
        ProblemContext.mergeLocation(
            ProblemContext.ProblemLocation.range("A1:B2"),
            ProblemContext.ProblemLocation.address("B4")));
    assertEquals(
        new ProblemContext.ProblemLocation.FormulaCell("Budget", "B4", "SUM(B2:B3)"),
        ProblemContext.mergeLocation(
            ProblemContext.ProblemLocation.range("A1:B2"),
            ProblemContext.ProblemLocation.formulaCell("Budget", "B4", "SUM(B2:B3)")));
    assertEquals(
        ProblemContext.ProblemLocation.cell("Budget", "B4"),
        ProblemContext.mergeLocation(
            ProblemContext.ProblemLocation.range("A1:B2"),
            ProblemContext.ProblemLocation.cell("Budget", "B4")));
    assertEquals(
        ProblemContext.ProblemLocation.namedRange("BudgetTotal"),
        ProblemContext.mergeLocation(
            ProblemContext.ProblemLocation.range("A1:B2"),
            ProblemContext.ProblemLocation.namedRange("BudgetTotal")));
    assertEquals(
        new ProblemContext.ProblemLocation.Range("Budget", "C1:D4"),
        ProblemContext.mergeLocation(
            ProblemContext.ProblemLocation.range("A1:B2"),
            new ProblemContext.ProblemLocation.Range("Budget", "C1:D4")));
    assertEquals(
        new ProblemContext.ProblemLocation.SheetNamedRange("Budget", "BudgetTotal"),
        ProblemContext.mergeLocation(
            ProblemContext.ProblemLocation.range("A1:B2"),
            new ProblemContext.ProblemLocation.SheetNamedRange("Budget", "BudgetTotal")));
    assertEquals(
        "field must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> ProblemContextSupport.requireNonBlank(" ", "field"))
            .getMessage());
    assertEquals(
        "field must not be null",
        assertThrows(
                NullPointerException.class,
                () -> ProblemContextSupport.requireNonBlank(null, "field"))
            .getMessage());
  }
}
