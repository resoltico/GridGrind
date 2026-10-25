package dev.erst.gridgrind.contract.step;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.query.WorkbookIntrospectionQuery;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.RangeSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
import dev.erst.gridgrind.contract.source.TextSourceInput;
import org.junit.jupiter.api.Test;

/** Verifies that step construction retains bindable fragments for static contract validation. */
class WorkbookStepValidationTest {
  @Test
  void keepsKnownButIncompatibleOperationTargetPairsBound() {
    CellMutationAction.SetCell action =
        new CellMutationAction.SetCell(new CellInput.Text(TextSourceInput.inline("Owner")));

    assertDoesNotThrow(
        () -> new MutationStep("set-owner", new RangeSelector.ByRange("Budget", "A1:B2"), action));
    assertEquals(
        "SET_CELL requires target type CELL_BY_ADDRESS or TABLE_CELL_BY_COLUMN_NAME but got RANGE_BY_RANGE",
        WorkbookOperationContracts.targetViolation(
                action, new RangeSelector.ByRange("Budget", "A1:B2"))
            .orElseThrow());
  }

  @Test
  void acceptsCompatibleOperationTargetPairsThroughTheSameContract() {
    CellMutationAction.SetCell action =
        new CellMutationAction.SetCell(new CellInput.Text(TextSourceInput.inline("Owner")));

    assertTrue(
        WorkbookOperationContracts.targetViolation(
                action, new CellSelector.ByAddress("Budget", "A1"))
            .isEmpty());
    assertTrue(
        WorkbookOperationContracts.targetViolation(
                new WorkbookIntrospectionQuery.GetWorkbookSummary(), new WorkbookSelector.Current())
            .isEmpty());
  }
}
