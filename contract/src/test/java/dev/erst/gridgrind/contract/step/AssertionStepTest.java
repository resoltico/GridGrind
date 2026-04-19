package dev.erst.gridgrind.contract.step;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.assertion.Assertion;
import dev.erst.gridgrind.contract.assertion.ExpectedCellValue;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
import org.junit.jupiter.api.Test;

/** Direct coverage for assertion-step validation and diagnostics. */
class AssertionStepTest {
  @Test
  void constructsAssertionStepsAndRejectsIncompatibleTargets() {
    Assertion assertion = new Assertion.CellValue(new ExpectedCellValue.Text("Owner"));
    AssertionStep step =
        new AssertionStep("assert-owner", new CellSelector.ByAddress("Budget", "A1"), assertion);

    assertEquals("assert-owner", step.stepId());
    assertEquals("ASSERTION", step.stepKind());
    assertEquals(assertion, step.assertion());

    IllegalArgumentException wrongTarget =
        assertThrows(
            IllegalArgumentException.class,
            () -> new AssertionStep("bad", new WorkbookSelector.Current(), assertion));
    assertEquals(
        "EXPECT_CELL_VALUE requires target type ByAddress, ByAddresses or ByColumnName but got Current",
        wrongTarget.getMessage());
  }
}
