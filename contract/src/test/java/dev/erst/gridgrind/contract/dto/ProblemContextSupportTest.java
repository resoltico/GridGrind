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
        new ProblemContextWorkbookSurfaces.ProblemLocation.Cell("Budget", "B4"),
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.sheet("Budget"),
            ProblemContextWorkbookSurfaces.ProblemLocation.address("B4")));
    assertEquals(
        new ProblemContextWorkbookSurfaces.ProblemLocation.Range("Budget", "A1:B2"),
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.sheet("Budget"),
            ProblemContextWorkbookSurfaces.ProblemLocation.range("A1:B2")));
    assertEquals(
        new ProblemContextWorkbookSurfaces.ProblemLocation.FormulaCell(
            "Budget", "B4", "SUM(B2:B3)"),
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.sheet("Budget"),
            ProblemContextWorkbookSurfaces.ProblemLocation.formulaCell(
                "Ignored", "B4", "SUM(B2:B3)")));
    assertEquals(
        new ProblemContextWorkbookSurfaces.ProblemLocation.FormulaCell(
            "Budget", "B4", "SUM(B2:B3)"),
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.cell("Budget", "B4"),
            ProblemContextWorkbookSurfaces.ProblemLocation.formulaCell(
                "Archive", "C9", "SUM(B2:B3)")));
    assertEquals(
        new ProblemContextWorkbookSurfaces.ProblemLocation.Cell("Budget", "B4"),
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.address("B4"),
            ProblemContextWorkbookSurfaces.ProblemLocation.sheet("Budget")));
    assertEquals(
        new ProblemContextWorkbookSurfaces.ProblemLocation.FormulaCell(
            "Budget", "B4", "SUM(B2:B3)"),
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.address("B4"),
            ProblemContextWorkbookSurfaces.ProblemLocation.formulaCell(
                "Budget", "C5", "SUM(B2:B3)")));
    assertEquals(
        new ProblemContextWorkbookSurfaces.ProblemLocation.Range("Budget", "A1:B2"),
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.range("A1:B2"),
            ProblemContextWorkbookSurfaces.ProblemLocation.sheet("Budget")));
    assertEquals(
        new ProblemContextWorkbookSurfaces.ProblemLocation.SheetNamedRange("Budget", "BudgetTotal"),
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.range("A1:B2"),
            ProblemContextWorkbookSurfaces.ProblemLocation.namedRange("Budget", "BudgetTotal")));
    assertEquals(
        ProblemContextWorkbookSurfaces.ProblemLocation.address("B4"),
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.unknown(),
            ProblemContextWorkbookSurfaces.ProblemLocation.address("B4")));
    assertEquals(
        ProblemContextWorkbookSurfaces.ProblemLocation.sheet("Budget"),
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.sheet("Budget"),
            ProblemContextWorkbookSurfaces.ProblemLocation.sheet("Archive")));
    assertEquals(
        new ProblemContextWorkbookSurfaces.ProblemLocation.Cell("Budget", "C7"),
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.sheet("Budget"),
            ProblemContextWorkbookSurfaces.ProblemLocation.cell("Archive", "C7")));
    assertEquals(
        ProblemContextWorkbookSurfaces.ProblemLocation.namedRange("BudgetTotal"),
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.sheet("Budget"),
            ProblemContextWorkbookSurfaces.ProblemLocation.namedRange("BudgetTotal")));
    assertEquals(
        new ProblemContextWorkbookSurfaces.ProblemLocation.Range("Budget", "C1:D4"),
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.sheet("Budget"),
            new ProblemContextWorkbookSurfaces.ProblemLocation.Range("Archive", "C1:D4")));
    assertEquals(
        new ProblemContextWorkbookSurfaces.ProblemLocation.SheetNamedRange("Ops", "BudgetTotal"),
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.sheet("Budget"),
            new ProblemContextWorkbookSurfaces.ProblemLocation.SheetNamedRange(
                "Ops", "BudgetTotal")));
  }

  @Test
  void mergeLocationPreservesSpecificCurrentLocationsAndRejectsBlankValues() {
    ProblemContextWorkbookSurfaces.ProblemLocation current =
        ProblemContextWorkbookSurfaces.ProblemLocation.formulaCell("Budget", "B4", "SUM(B2:B3)");

    assertEquals(
        current,
        ProblemContext.mergeLocation(
            current, ProblemContextWorkbookSurfaces.ProblemLocation.sheet("Archive")));
    assertInstanceOf(
        ProblemContextWorkbookSurfaces.ProblemLocation.NamedRange.class,
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.sheet("Budget"),
            ProblemContextWorkbookSurfaces.ProblemLocation.namedRange("BudgetTotal")));
    assertEquals(
        ProblemContextWorkbookSurfaces.ProblemLocation.namedRange("Budget", "BudgetTotal"),
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.namedRange("Budget", "BudgetTotal"),
            ProblemContextWorkbookSurfaces.ProblemLocation.cell("Budget", "B4")));
    assertEquals(
        ProblemContextWorkbookSurfaces.ProblemLocation.namedRange("BudgetTotal"),
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.namedRange("BudgetTotal"),
            ProblemContextWorkbookSurfaces.ProblemLocation.cell("Budget", "B4")));
    assertEquals(
        new ProblemContextWorkbookSurfaces.ProblemLocation.Range("Budget", "A1:B2"),
        ProblemContext.mergeLocation(
            new ProblemContextWorkbookSurfaces.ProblemLocation.Range("Budget", "A1:B2"),
            ProblemContextWorkbookSurfaces.ProblemLocation.namedRange("BudgetTotal")));
    assertEquals(
        new ProblemContextWorkbookSurfaces.ProblemLocation.Cell("Budget", "B4"),
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.cell("Budget", "B4"),
            ProblemContextWorkbookSurfaces.ProblemLocation.sheet("Archive")));
    assertEquals(
        ProblemContextWorkbookSurfaces.ProblemLocation.address("B4"),
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.address("B4"),
            ProblemContextWorkbookSurfaces.ProblemLocation.range("A1:B2")));
    assertEquals(
        ProblemContextWorkbookSurfaces.ProblemLocation.range("A1:B2"),
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.range("A1:B2"),
            ProblemContextWorkbookSurfaces.ProblemLocation.address("B4")));
    assertEquals(
        new ProblemContextWorkbookSurfaces.ProblemLocation.FormulaCell(
            "Budget", "B4", "SUM(B2:B3)"),
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.range("A1:B2"),
            ProblemContextWorkbookSurfaces.ProblemLocation.formulaCell(
                "Budget", "B4", "SUM(B2:B3)")));
    assertEquals(
        ProblemContextWorkbookSurfaces.ProblemLocation.cell("Budget", "B4"),
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.range("A1:B2"),
            ProblemContextWorkbookSurfaces.ProblemLocation.cell("Budget", "B4")));
    assertEquals(
        ProblemContextWorkbookSurfaces.ProblemLocation.namedRange("BudgetTotal"),
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.range("A1:B2"),
            ProblemContextWorkbookSurfaces.ProblemLocation.namedRange("BudgetTotal")));
    assertEquals(
        new ProblemContextWorkbookSurfaces.ProblemLocation.Range("Budget", "C1:D4"),
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.range("A1:B2"),
            new ProblemContextWorkbookSurfaces.ProblemLocation.Range("Budget", "C1:D4")));
    assertEquals(
        new ProblemContextWorkbookSurfaces.ProblemLocation.SheetNamedRange("Budget", "BudgetTotal"),
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.range("A1:B2"),
            new ProblemContextWorkbookSurfaces.ProblemLocation.SheetNamedRange(
                "Budget", "BudgetTotal")));
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
