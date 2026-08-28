package dev.erst.gridgrind.engine.runtime;

import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.inspect;
import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.inspections;
import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.mutate;
import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.mutations;
import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.request;
import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.text;
import static dev.erst.gridgrind.engine.runtime.ExecutorTestPlanSupport.textCell;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.action.CellMutationAction;
import dev.erst.gridgrind.contract.action.StructuredMutationAction;
import dev.erst.gridgrind.contract.action.WorkbookMutationAction;
import dev.erst.gridgrind.contract.assertion.*;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.TableInput;
import dev.erst.gridgrind.contract.dto.TableStyleInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.dto.WorkbookResult;
import dev.erst.gridgrind.contract.query.*;
import dev.erst.gridgrind.contract.query.SheetInspectionResult;
import dev.erst.gridgrind.contract.selector.TableCellSelector;
import dev.erst.gridgrind.contract.selector.TableRowSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Coverage for semantic selector execution and zero-match behavior. */
class SemanticSelectorExecutionCoverageTest {
  @Test
  void tableCellInspectionReturnsEmptyResultWhenKeySelectorMatchesNoRow() {
    WorkbookResult.Success success =
        success(
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    budgetTableMutations(),
                    inspections(
                        inspect(
                            "inspect-missing-table-cell",
                            missingAmountCellTarget(),
                            new SheetIntrospectionQuery.GetCells())))));

    SheetInspectionResult.CellsResult cellsResult =
        assertInstanceOf(SheetInspectionResult.CellsResult.class, success.inspections().getFirst());
    assertEquals("Budget", cellsResult.sheetName());
    assertEquals(List.of(), cellsResult.cells());
  }

  @Test
  void tableCellAssertionsFailStructurallyWhenKeySelectorMatchesNoRow() {
    WorkbookResult.Failure failure =
        failure(
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    budgetTableMutations(),
                    ExecutorTestPlanSupport.assertions(
                        ExecutorTestPlanSupport.assertThat(
                            "assert-missing-table-cell",
                            missingAmountCellTarget(),
                            new CellAssertion.CellValue(
                                new dev.erst.gridgrind.contract.dto.CellScalarValue.NumberValue(
                                    999.0)))),
                    List.of())));

    assertEquals(GridGrindProblemCode.ASSERTION_FAILED, failure.problem().code());
    assertTrue(
        failure.problem().message().contains("EXPECT_CELL_VALUE resolved no matching cells"));
    assertEquals("assert-missing-table-cell", failedAssertion(failure).stepId());
    SheetInspectionResult.CellsResult cellsResult =
        assertInstanceOf(
            SheetInspectionResult.CellsResult.class,
            failedAssertion(failure).observations().getFirst());
    assertEquals(List.of(), cellsResult.cells());
  }

  private static dev.erst.gridgrind.contract.assertion.AssertionFailure failedAssertion(
      WorkbookResult.Failure failure) {
    return assertInstanceOf(AssertionResult.Failed.class, failure.assertions().getFirst())
        .failure();
  }

  @Test
  void semanticResolverRejectsDuplicateKeyMatchesInsteadOfGuessingOneRow() {
    WorkbookResult.Failure failure =
        failure(
            ExecutionContextFixtureSupport.execute(
                new DefaultGridGrindRequestExecutor(),
                request(
                    new WorkbookPlan.WorkbookSource.New(),
                    new WorkbookPlan.WorkbookPersistence.None(),
                    duplicateBudgetTableMutations(),
                    inspections(
                        inspect(
                            "inspect-duplicate-table-cell",
                            new TableCellSelector.ByColumnName(
                                new TableRowSelector.ByKeyCell(
                                    new TableSelector.ByName("BudgetTable"),
                                    "Item",
                                    textCell("Hosting")),
                                "Amount"),
                            new SheetIntrospectionQuery.GetCells())))));

    assertEquals(GridGrindProblemCode.INVALID_REQUEST, failure.problem().code());
    assertTrue(failure.problem().message().contains("matched more than one row"));
  }

  private static TableCellSelector.ByColumnName missingAmountCellTarget() {
    return new TableCellSelector.ByColumnName(
        new TableRowSelector.ByKeyCell(
            new TableSelector.ByName("BudgetTable"), "Item", textCell("Missing")),
        "Amount");
  }

  private static List<ExecutorTestPlanSupport.PendingMutation> budgetTableMutations() {
    return mutations(
        mutate(
            new dev.erst.gridgrind.contract.selector.SheetSelector.ByName("Budget"),
            new WorkbookMutationAction.EnsureSheet()),
        mutate(
            new dev.erst.gridgrind.contract.selector.RangeSelector.ByRange("Budget", "A1:B3"),
            new CellMutationAction.SetRange(
                new dev.erst.gridgrind.contract.dto.CellGridInput.Typed(
                    List.of(
                        List.of(
                            new CellInput.Text(text("Item")), new CellInput.Text(text("Amount"))),
                        List.of(
                            new CellInput.Text(text("Hosting")), new CellInput.NumberValue(100.0)),
                        List.of(
                            new CellInput.Text(text("Travel")),
                            new CellInput.NumberValue(50.0)))))),
        mutate(
            new StructuredMutationAction.SetTable(
                TableInput.withDefaultMetadata(
                    "BudgetTable", "Budget", "A1:B3", false, new TableStyleInput.None()))));
  }

  private static List<ExecutorTestPlanSupport.PendingMutation> duplicateBudgetTableMutations() {
    return mutations(
        mutate(
            new dev.erst.gridgrind.contract.selector.SheetSelector.ByName("Budget"),
            new WorkbookMutationAction.EnsureSheet()),
        mutate(
            new dev.erst.gridgrind.contract.selector.RangeSelector.ByRange("Budget", "A1:B4"),
            new CellMutationAction.SetRange(
                new dev.erst.gridgrind.contract.dto.CellGridInput.Typed(
                    List.of(
                        List.of(
                            new CellInput.Text(text("Item")), new CellInput.Text(text("Amount"))),
                        List.of(
                            new CellInput.Text(text("Hosting")), new CellInput.NumberValue(100.0)),
                        List.of(
                            new CellInput.Text(text("Hosting")), new CellInput.NumberValue(125.0)),
                        List.of(
                            new CellInput.Text(text("Travel")),
                            new CellInput.NumberValue(50.0)))))),
        mutate(
            new StructuredMutationAction.SetTable(
                TableInput.withDefaultMetadata(
                    "BudgetTable", "Budget", "A1:B4", false, new TableStyleInput.None()))));
  }

  private static WorkbookResult.Success success(WorkbookResult response) {
    return assertInstanceOf(WorkbookResult.Success.class, response);
  }

  private static WorkbookResult.Failure failure(WorkbookResult response) {
    return assertInstanceOf(WorkbookResult.Failure.class, response);
  }
}
