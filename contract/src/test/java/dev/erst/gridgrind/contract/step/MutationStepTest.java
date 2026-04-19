package dev.erst.gridgrind.contract.step;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import org.junit.jupiter.api.Test;

/** Tests for ordered mutation step validation and target compatibility. */
class MutationStepTest {
  @Test
  void acceptsCompatibleTargetAndActionPairs() {
    MutationStep ensureSheet =
        new MutationStep(
            "ensure-budget", new SheetSelector.ByName("Budget"), new MutationAction.EnsureSheet());
    MutationStep setCell =
        new MutationStep(
            "set-owner",
            new CellSelector.ByAddress("Budget", "A1"),
            new MutationAction.SetCell(new CellInput.Text(text("Owner"))));
    MutationStep clearWorkbookProtection =
        new MutationStep(
            "clear-book-protection",
            new WorkbookSelector.Current(),
            new MutationAction.ClearWorkbookProtection());

    assertEquals("MUTATION", ensureSheet.stepKind());
    assertEquals("ENSURE_SHEET", ensureSheet.action().actionType());
    assertEquals("A1", ((CellSelector.ByAddress) setCell.target()).address());
    assertEquals("CLEAR_WORKBOOK_PROTECTION", clearWorkbookProtection.action().actionType());
  }

  @Test
  void rejectsIncompatibleTargetsOrBlankStepIds() {
    assertThrows(
        IllegalArgumentException.class,
        () ->
            new MutationStep(
                " ", new SheetSelector.ByName("Budget"), new MutationAction.EnsureSheet()));
    IllegalArgumentException incompatibleTargetFailure =
        assertThrows(
            IllegalArgumentException.class,
            () ->
                new MutationStep(
                    "bad-target",
                    new RangeSelector.ByRange("Budget", "A1:B2"),
                    new MutationAction.SetCell(new CellInput.Text(text("Owner")))));

    assertEquals(
        "SET_CELL requires target type ByAddress or ByColumnName but got ByRange",
        incompatibleTargetFailure.getMessage());
  }

  private static TextSourceInput text(String value) {
    return TextSourceInput.inline(value);
  }
}
