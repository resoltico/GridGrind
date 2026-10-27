package dev.erst.gridgrind.core.problem;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.vocabulary.SheetLabel;
import dev.erst.gridgrind.core.vocabulary.TextValidation;
import org.junit.jupiter.api.Test;

/** Direct coverage for location-merging helpers behind the typed problem-context model. */
class ProblemContextSupportTest {
  @Test
  void mergeLocationComposesSheetCellRangeNamedRangeAndFormulaVariants() {
    assertEquals(
        new ProblemContextWorkbookSurfaces.ProblemLocation.Cell(new SheetLabel("Budget"), "B4"),
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.sheet(new SheetLabel("Budget")),
            ProblemContextWorkbookSurfaces.ProblemLocation.address("B4")));
    assertEquals(
        new ProblemContextWorkbookSurfaces.ProblemLocation.Range(new SheetLabel("Budget"), "A1:B2"),
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.sheet(new SheetLabel("Budget")),
            ProblemContextWorkbookSurfaces.ProblemLocation.range("A1:B2")));
    assertEquals(
        new ProblemContextWorkbookSurfaces.ProblemLocation.FormulaCell(
            new SheetLabel("Budget"), "B4", "SUM(B2:B3)"),
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.sheet(new SheetLabel("Budget")),
            ProblemContextWorkbookSurfaces.ProblemLocation.formulaCell(
                new SheetLabel("Ignored"), "B4", "SUM(B2:B3)")));
    assertEquals(
        new ProblemContextWorkbookSurfaces.ProblemLocation.FormulaCell(
            new SheetLabel("Budget"), "B4", "SUM(B2:B3)"),
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.cell(new SheetLabel("Budget"), "B4"),
            ProblemContextWorkbookSurfaces.ProblemLocation.formulaCell(
                new SheetLabel("Archive"), "C9", "SUM(B2:B3)")));
    assertEquals(
        new ProblemContextWorkbookSurfaces.ProblemLocation.Cell(new SheetLabel("Budget"), "B4"),
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.address("B4"),
            ProblemContextWorkbookSurfaces.ProblemLocation.sheet(new SheetLabel("Budget"))));
    assertEquals(
        new ProblemContextWorkbookSurfaces.ProblemLocation.FormulaCell(
            new SheetLabel("Budget"), "B4", "SUM(B2:B3)"),
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.address("B4"),
            ProblemContextWorkbookSurfaces.ProblemLocation.formulaCell(
                new SheetLabel("Budget"), "C5", "SUM(B2:B3)")));
    assertEquals(
        new ProblemContextWorkbookSurfaces.ProblemLocation.Range(new SheetLabel("Budget"), "A1:B2"),
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.range("A1:B2"),
            ProblemContextWorkbookSurfaces.ProblemLocation.sheet(new SheetLabel("Budget"))));
    assertEquals(
        new ProblemContextWorkbookSurfaces.ProblemLocation.SheetNamedRange(
            new SheetLabel("Budget"), "BudgetTotal"),
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.range("A1:B2"),
            ProblemContextWorkbookSurfaces.ProblemLocation.namedRange(
                new SheetLabel("Budget"), "BudgetTotal")));
    assertEquals(
        ProblemContextWorkbookSurfaces.ProblemLocation.address("B4"),
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.unknown(),
            ProblemContextWorkbookSurfaces.ProblemLocation.address("B4")));
    assertEquals(
        ProblemContextWorkbookSurfaces.ProblemLocation.sheet(new SheetLabel("Budget")),
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.sheet(new SheetLabel("Budget")),
            ProblemContextWorkbookSurfaces.ProblemLocation.sheet(new SheetLabel("Archive"))));
    assertEquals(
        new ProblemContextWorkbookSurfaces.ProblemLocation.Cell(new SheetLabel("Budget"), "C7"),
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.sheet(new SheetLabel("Budget")),
            ProblemContextWorkbookSurfaces.ProblemLocation.cell(new SheetLabel("Archive"), "C7")));
    assertEquals(
        ProblemContextWorkbookSurfaces.ProblemLocation.namedRange("BudgetTotal"),
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.sheet(new SheetLabel("Budget")),
            ProblemContextWorkbookSurfaces.ProblemLocation.namedRange("BudgetTotal")));
    assertEquals(
        new ProblemContextWorkbookSurfaces.ProblemLocation.Range(new SheetLabel("Budget"), "C1:D4"),
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.sheet(new SheetLabel("Budget")),
            new ProblemContextWorkbookSurfaces.ProblemLocation.Range(
                new SheetLabel("Archive"), "C1:D4")));
    assertEquals(
        new ProblemContextWorkbookSurfaces.ProblemLocation.SheetNamedRange(
            new SheetLabel("Ops"), "BudgetTotal"),
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.sheet(new SheetLabel("Budget")),
            new ProblemContextWorkbookSurfaces.ProblemLocation.SheetNamedRange(
                new SheetLabel("Ops"), "BudgetTotal")));
  }

  @Test
  void mergeLocationPreservesSpecificCurrentLocationsAndRejectsBlankValues() {
    ProblemContextWorkbookSurfaces.ProblemLocation current =
        ProblemContextWorkbookSurfaces.ProblemLocation.formulaCell(
            new SheetLabel("Budget"), "B4", "SUM(B2:B3)");

    assertEquals(
        current,
        ProblemContext.mergeLocation(
            current,
            ProblemContextWorkbookSurfaces.ProblemLocation.sheet(new SheetLabel("Archive"))));
    assertInstanceOf(
        ProblemContextWorkbookSurfaces.ProblemLocation.NamedRange.class,
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.sheet(new SheetLabel("Budget")),
            ProblemContextWorkbookSurfaces.ProblemLocation.namedRange("BudgetTotal")));
    assertEquals(
        ProblemContextWorkbookSurfaces.ProblemLocation.namedRange(
            new SheetLabel("Budget"), "BudgetTotal"),
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.namedRange(
                new SheetLabel("Budget"), "BudgetTotal"),
            ProblemContextWorkbookSurfaces.ProblemLocation.cell(new SheetLabel("Budget"), "B4")));
    assertEquals(
        ProblemContextWorkbookSurfaces.ProblemLocation.namedRange("BudgetTotal"),
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.namedRange("BudgetTotal"),
            ProblemContextWorkbookSurfaces.ProblemLocation.cell(new SheetLabel("Budget"), "B4")));
    assertEquals(
        new ProblemContextWorkbookSurfaces.ProblemLocation.Range(new SheetLabel("Budget"), "A1:B2"),
        ProblemContext.mergeLocation(
            new ProblemContextWorkbookSurfaces.ProblemLocation.Range(
                new SheetLabel("Budget"), "A1:B2"),
            ProblemContextWorkbookSurfaces.ProblemLocation.namedRange("BudgetTotal")));
    assertEquals(
        new ProblemContextWorkbookSurfaces.ProblemLocation.Cell(new SheetLabel("Budget"), "B4"),
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.cell(new SheetLabel("Budget"), "B4"),
            ProblemContextWorkbookSurfaces.ProblemLocation.sheet(new SheetLabel("Archive"))));
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
            new SheetLabel("Budget"), "B4", "SUM(B2:B3)"),
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.range("A1:B2"),
            ProblemContextWorkbookSurfaces.ProblemLocation.formulaCell(
                new SheetLabel("Budget"), "B4", "SUM(B2:B3)")));
    assertEquals(
        ProblemContextWorkbookSurfaces.ProblemLocation.cell(new SheetLabel("Budget"), "B4"),
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.range("A1:B2"),
            ProblemContextWorkbookSurfaces.ProblemLocation.cell(new SheetLabel("Budget"), "B4")));
    assertEquals(
        ProblemContextWorkbookSurfaces.ProblemLocation.namedRange("BudgetTotal"),
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.range("A1:B2"),
            ProblemContextWorkbookSurfaces.ProblemLocation.namedRange("BudgetTotal")));
    assertEquals(
        new ProblemContextWorkbookSurfaces.ProblemLocation.Range(new SheetLabel("Budget"), "C1:D4"),
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.range("A1:B2"),
            new ProblemContextWorkbookSurfaces.ProblemLocation.Range(
                new SheetLabel("Budget"), "C1:D4")));
    assertEquals(
        new ProblemContextWorkbookSurfaces.ProblemLocation.SheetNamedRange(
            new SheetLabel("Budget"), "BudgetTotal"),
        ProblemContext.mergeLocation(
            ProblemContextWorkbookSurfaces.ProblemLocation.range("A1:B2"),
            new ProblemContextWorkbookSurfaces.ProblemLocation.SheetNamedRange(
                new SheetLabel("Budget"), "BudgetTotal")));
    assertEquals(
        "field must not be blank",
        assertThrows(
                IllegalArgumentException.class, () -> TextValidation.requireNonBlank(" ", "field"))
            .getMessage());
    assertEquals(
        "field must not be null",
        assertThrows(
                NullPointerException.class, () -> TextValidation.requireNonBlank(null, "field"))
            .getMessage());
  }
}
