package dev.erst.gridgrind.contract.step;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.contract.assertion.*;
import dev.erst.gridgrind.contract.selector.CellSelector;
import dev.erst.gridgrind.contract.selector.WorkbookSelector;
import org.junit.jupiter.api.Test;

/** Direct coverage for assertion-step validation and diagnostics. */
class AssertionStepTest {
  @Test
  void constructsAssertionStepsAndKeepsIncompatibleTargetsForStaticValidation() {
    Assertion assertion =
        new CellAssertion.CellValue(
            new dev.erst.gridgrind.contract.dto.CellScalarValue.Text("Owner"));
    AssertionStep step =
        new AssertionStep("assert-owner", new CellSelector.ByAddress("Budget", "A1"), assertion);

    assertEquals("assert-owner", step.stepId());
    assertEquals("ASSERTION", step.stepKind());
    assertEquals(assertion, step.assertion());

    assertDoesNotThrow(() -> new AssertionStep("bad", new WorkbookSelector.Current(), assertion));
  }
}
