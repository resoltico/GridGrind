package dev.erst.gridgrind.cli.examples;

import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.contract.selector.TableSelector;
import org.junit.jupiter.api.Test;

/** Focused tests for package-private example-plan helpers. */
class ExamplePlanSupportTest {
  @Test
  void tableHelperBuildsSheetOwnedTableSelectors() {
    TableSelector.ByNameOnSheet selector = ExamplePlanSupport.table("BudgetTable", "Budget");

    assertEquals("BudgetTable", selector.name());
    assertEquals("Budget", selector.sheetName());
  }
}
