package dev.erst.gridgrind.core.plan;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.model.value.CellInput;
import dev.erst.gridgrind.core.source.TextSourceInput;
import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import org.junit.jupiter.api.Test;

/** Tests for ordered mutation step validation and target compatibility. */
class MutationStepTest {
  @Test
  void acceptsCompatibleTargetAndActionPairs() {
    MutationStep ensureSheet =
        new MutationStep(
            new StepId("ensure-budget"),
            new SheetSelector.ByName(new SheetName("Budget")),
            new WorkbookMutationAction.EnsureSheet());
    MutationStep setCell =
        new MutationStep(
            new StepId("set-owner"),
            new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
            new CellMutationAction.SetCell(new CellInput.Text(text("Owner"))));
    MutationStep clearWorkbookProtection =
        new MutationStep(
            new StepId("clear-book-protection"),
            new WorkbookSelector.Current(),
            new WorkbookMutationAction.ClearWorkbookProtection());

    assertEquals("MUTATION", ensureSheet.stepKind());
    assertEquals("ENSURE_SHEET", ensureSheet.action().actionType());
    assertEquals("A1", ((CellSelector.ByAddress) setCell.target()).address().value());
    assertEquals("CLEAR_WORKBOOK_PROTECTION", clearWorkbookProtection.action().actionType());
  }

  @Test
  void keepsIncompatibleTargetsForTheStaticContractPhase() {
    assertDoesNotThrow(
        () ->
            new MutationStep(
                new StepId("bad-target"),
                new RangeSelector.ByRange(new SheetName("Budget"), new RangeRef("A1:B2")),
                new CellMutationAction.SetCell(new CellInput.Text(text("Owner")))));
  }

  @Test
  void rejectsBlankAndMalformedStepIdentifiersAtTheRecordBoundary() {
    assertEquals(
        "stepId must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new MutationStep(
                        new StepId(" "),
                        new SheetSelector.ByName(new SheetName("Budget")),
                        new WorkbookMutationAction.EnsureSheet()))
            .getMessage());
    assertEquals(
        "stepId must match [A-Za-z0-9._-]+",
        assertThrows(
                IllegalArgumentException.class,
                () ->
                    new MutationStep(
                        new StepId("not valid"),
                        new SheetSelector.ByName(new SheetName("Budget")),
                        new WorkbookMutationAction.EnsureSheet()))
            .getMessage());
  }

  private static TextSourceInput text(String value) {
    return TextSourceInput.inline(value);
  }
}
