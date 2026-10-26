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

import dev.erst.gridgrind.core.model.table.TableSpec;
import dev.erst.gridgrind.core.model.table.TableStyleSpec;
import dev.erst.gridgrind.core.model.value.CellValueSpec;
import dev.erst.gridgrind.core.plan.*;
import dev.erst.gridgrind.core.plan.CellMutationAction;
import dev.erst.gridgrind.core.plan.SheetInspectionResult;
import dev.erst.gridgrind.core.plan.StructuredMutationAction;
import dev.erst.gridgrind.core.plan.TableCellSelector;
import dev.erst.gridgrind.core.plan.TableRowSelector;
import dev.erst.gridgrind.core.plan.TableSelector;
import dev.erst.gridgrind.core.plan.WorkbookMutationAction;
import dev.erst.gridgrind.core.plan.WorkbookPlan;
import dev.erst.gridgrind.core.problem.GridGrindProblemCode;
import dev.erst.gridgrind.core.result.WorkbookResult;
import dev.erst.gridgrind.core.vocabulary.RangeRef;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.TableName;
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
    assertEquals("Budget", cellsResult.sheetName().value());
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
                                new dev.erst.gridgrind.core.model.value.CellScalarValue.NumberValue(
                                    999.0)))),
                    List.of())));

    assertEquals(GridGrindProblemCode.ASSERTION_FAILED, failure.problem().code());
    assertTrue(
        failure.problem().message().contains("EXPECT_CELL_VALUE resolved no matching cells"));
    assertEquals("assert-missing-table-cell", failedAssertion(failure).stepId().value());
    SheetInspectionResult.CellsResult cellsResult =
        assertInstanceOf(
            SheetInspectionResult.CellsResult.class,
            failedAssertion(failure).observations().getFirst());
    assertEquals(List.of(), cellsResult.cells());
  }

  private static dev.erst.gridgrind.core.plan.AssertionFailure failedAssertion(
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
                                    new TableSelector.ByName(new TableName("BudgetTable")),
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
            new TableSelector.ByName(new TableName("BudgetTable")), "Item", textCell("Missing")),
        "Amount");
  }

  private static List<ExecutorTestPlanSupport.PendingMutation> budgetTableMutations() {
    return mutations(
        mutate(
            new dev.erst.gridgrind.core.plan.SheetSelector.ByName(new SheetName("Budget")),
            new WorkbookMutationAction.EnsureSheet()),
        mutate(
            new dev.erst.gridgrind.core.plan.RangeSelector.ByRange(
                new SheetName("Budget"), new RangeRef("A1:B3")),
            new CellMutationAction.SetRange(
                new dev.erst.gridgrind.core.model.value.CellGridSpec.Typed(
                    List.of(
                        List.of(
                            new CellValueSpec.Text(text("Item")),
                            new CellValueSpec.Text(text("Amount"))),
                        List.of(
                            new CellValueSpec.Text(text("Hosting")),
                            new CellValueSpec.NumberValue(100.0)),
                        List.of(
                            new CellValueSpec.Text(text("Travel")),
                            new CellValueSpec.NumberValue(50.0)))))),
        mutate(
            new StructuredMutationAction.SetTable(
                TableSpec.withDefaultMetadata(
                    new TableName("BudgetTable"),
                    new SheetName("Budget"),
                    new RangeRef("A1:B3"),
                    false,
                    new TableStyleSpec.None()))));
  }

  private static List<ExecutorTestPlanSupport.PendingMutation> duplicateBudgetTableMutations() {
    return mutations(
        mutate(
            new dev.erst.gridgrind.core.plan.SheetSelector.ByName(new SheetName("Budget")),
            new WorkbookMutationAction.EnsureSheet()),
        mutate(
            new dev.erst.gridgrind.core.plan.RangeSelector.ByRange(
                new SheetName("Budget"), new RangeRef("A1:B4")),
            new CellMutationAction.SetRange(
                new dev.erst.gridgrind.core.model.value.CellGridSpec.Typed(
                    List.of(
                        List.of(
                            new CellValueSpec.Text(text("Item")),
                            new CellValueSpec.Text(text("Amount"))),
                        List.of(
                            new CellValueSpec.Text(text("Hosting")),
                            new CellValueSpec.NumberValue(100.0)),
                        List.of(
                            new CellValueSpec.Text(text("Hosting")),
                            new CellValueSpec.NumberValue(125.0)),
                        List.of(
                            new CellValueSpec.Text(text("Travel")),
                            new CellValueSpec.NumberValue(50.0)))))),
        mutate(
            new StructuredMutationAction.SetTable(
                TableSpec.withDefaultMetadata(
                    new TableName("BudgetTable"),
                    new SheetName("Budget"),
                    new RangeRef("A1:B4"),
                    false,
                    new TableStyleSpec.None()))));
  }

  private static WorkbookResult.Success success(WorkbookResult response) {
    return assertInstanceOf(WorkbookResult.Success.class, response);
  }

  private static WorkbookResult.Failure failure(WorkbookResult response) {
    return assertInstanceOf(WorkbookResult.Failure.class, response);
  }
}
