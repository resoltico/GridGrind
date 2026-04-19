package dev.erst.gridgrind.executor;

import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.inspect;
import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.inspections;
import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.mutate;
import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.mutations;
import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.request;
import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.text;
import static dev.erst.gridgrind.executor.ExecutorTestPlanSupport.textCell;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertTrue;

import dev.erst.gridgrind.contract.action.MutationAction;
import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.assertion.ExpectedCellValue;
import dev.erst.gridgrind.contract.dto.CellInput;
import dev.erst.gridgrind.contract.dto.GridGrindProblemCode;
import dev.erst.gridgrind.contract.dto.GridGrindResponse;
import dev.erst.gridgrind.contract.dto.TableInput;
import dev.erst.gridgrind.contract.dto.TableStyleInput;
import dev.erst.gridgrind.contract.dto.WorkbookPlan;
import dev.erst.gridgrind.contract.query.InspectionQuery;
import dev.erst.gridgrind.contract.query.InspectionResult;
import dev.erst.gridgrind.contract.selector.TableCellSelector;
import dev.erst.gridgrind.contract.selector.TableRowSelector;
import dev.erst.gridgrind.contract.selector.TableSelector;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Phase 8 coverage for semantic selector execution and zero-match behavior. */
class Phase8ExecutorCoverageTest {
  @Test
  void tableCellInspectionReturnsEmptyResultWhenKeySelectorMatchesNoRow() {
    GridGrindResponse.Success success =
        success(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        budgetTableMutations(),
                        inspections(
                            inspect(
                                "inspect-missing-table-cell",
                                missingAmountCellTarget(),
                                new InspectionQuery.GetCells())))));

    InspectionResult.CellsResult cellsResult =
        assertInstanceOf(InspectionResult.CellsResult.class, success.inspections().getFirst());
    assertEquals("Budget", cellsResult.sheetName());
    assertEquals(List.of(), cellsResult.cells());
  }

  @Test
  void tableCellAssertionsFailStructurallyWhenKeySelectorMatchesNoRow() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
                    request(
                        new WorkbookPlan.WorkbookSource.New(),
                        new WorkbookPlan.WorkbookPersistence.None(),
                        budgetTableMutations(),
                        ExecutorTestPlanSupport.assertions(
                            ExecutorTestPlanSupport.assertThat(
                                "assert-missing-table-cell",
                                missingAmountCellTarget(),
                                new Assertion.CellValue(
                                    new ExpectedCellValue.NumericValue(999.0)))),
                        List.of())));

    assertEquals(GridGrindProblemCode.ASSERTION_FAILED, failure.problem().code());
    assertTrue(
        failure.problem().message().contains("EXPECT_CELL_VALUE resolved no matching cells"));
    assertEquals("assert-missing-table-cell", failure.problem().assertionFailure().stepId());
    InspectionResult.CellsResult cellsResult =
        assertInstanceOf(
            InspectionResult.CellsResult.class,
            failure.problem().assertionFailure().observations().getFirst());
    assertEquals(List.of(), cellsResult.cells());
  }

  @Test
  void semanticResolverRejectsDuplicateKeyMatchesInsteadOfGuessingOneRow() {
    GridGrindResponse.Failure failure =
        failure(
            new DefaultGridGrindRequestExecutor()
                .execute(
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
                                new InspectionQuery.GetCells())))));

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
            new MutationAction.EnsureSheet()),
        mutate(
            new dev.erst.gridgrind.contract.selector.RangeSelector.ByRange("Budget", "A1:B3"),
            new MutationAction.SetRange(
                List.of(
                    List.of(new CellInput.Text(text("Item")), new CellInput.Text(text("Amount"))),
                    List.of(new CellInput.Text(text("Hosting")), new CellInput.Numeric(100.0)),
                    List.of(new CellInput.Text(text("Travel")), new CellInput.Numeric(50.0))))),
        mutate(
            new MutationAction.SetTable(
                new TableInput(
                    "BudgetTable", "Budget", "A1:B3", false, new TableStyleInput.None()))));
  }

  private static List<ExecutorTestPlanSupport.PendingMutation> duplicateBudgetTableMutations() {
    return mutations(
        mutate(
            new dev.erst.gridgrind.contract.selector.SheetSelector.ByName("Budget"),
            new MutationAction.EnsureSheet()),
        mutate(
            new dev.erst.gridgrind.contract.selector.RangeSelector.ByRange("Budget", "A1:B4"),
            new MutationAction.SetRange(
                List.of(
                    List.of(new CellInput.Text(text("Item")), new CellInput.Text(text("Amount"))),
                    List.of(new CellInput.Text(text("Hosting")), new CellInput.Numeric(100.0)),
                    List.of(new CellInput.Text(text("Hosting")), new CellInput.Numeric(125.0)),
                    List.of(new CellInput.Text(text("Travel")), new CellInput.Numeric(50.0))))),
        mutate(
            new MutationAction.SetTable(
                new TableInput(
                    "BudgetTable", "Budget", "A1:B4", false, new TableStyleInput.None()))));
  }

  private static GridGrindResponse.Success success(GridGrindResponse response) {
    return assertInstanceOf(GridGrindResponse.Success.class, response);
  }

  private static GridGrindResponse.Failure failure(GridGrindResponse response) {
    return assertInstanceOf(GridGrindResponse.Failure.class, response);
  }
}
