package dev.erst.gridgrind.core.semantics;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.core.plan.CellAssertion;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookIntrospectionQuery;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Verifies complete static semantics derivation for concrete operation record classes. */
class WorkbookOperationSemanticsFactoryTest {
  @Test
  void derivesMutationAndReadOnlySemanticsAndRejectsUnknownRecords() {
    assertTrue(
        WorkbookOperationSemanticsFactory.semanticsFor(CellMutationAction.ClearRange.class)
            .effects()
            .containsAll(List.of(OperationEffect.READ_WORKBOOK, OperationEffect.MUTATE_WORKBOOK)));
    assertEquals(
        List.of(
            new dev.erst.gridgrind.core.semantics.OperationPrecondition
                .ColumnEditsBeforeFormulaAuthoring()),
        WorkbookOperationSemanticsFactory.semanticsFor(WorkbookMutationAction.InsertColumns.class)
            .preconditions());
    assertEquals(
        List.of(OperationEffect.READ_WORKBOOK),
        WorkbookOperationSemanticsFactory.semanticsFor(CellAssertion.Value.class).effects());
    assertEquals(
        List.of(OperationEffect.READ_WORKBOOK),
        WorkbookOperationSemanticsFactory.semanticsFor(
                WorkbookIntrospectionQuery.GetWorkbookSummary.class)
            .effects());
    assertThrows(
        IllegalArgumentException.class,
        () -> WorkbookOperationSemanticsFactory.semanticsFor(UnsupportedOperation.class));
  }

  private record UnsupportedOperation() {}
}
