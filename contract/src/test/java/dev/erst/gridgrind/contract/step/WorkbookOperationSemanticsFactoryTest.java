package dev.erst.gridgrind.contract.step;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.assertion.CellAssertion;
import dev.erst.gridgrind.contract.catalog.OperationEffect;
import dev.erst.gridgrind.contract.query.WorkbookIntrospectionQuery;
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
            new dev.erst.gridgrind.contract.catalog.OperationPrecondition
                .ColumnEditsBeforeFormulaAuthoring()),
        WorkbookOperationSemanticsFactory.semanticsFor(WorkbookMutationAction.InsertColumns.class)
            .preconditions());
    assertEquals(
        List.of(OperationEffect.READ_WORKBOOK),
        WorkbookOperationSemanticsFactory.semanticsFor(CellAssertion.CellValue.class).effects());
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
