package dev.erst.gridgrind.core.plan;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;

import dev.erst.gridgrind.core.vocabulary.CellAddress;
import dev.erst.gridgrind.core.vocabulary.SheetName;
import dev.erst.gridgrind.core.vocabulary.StepId;
import org.junit.jupiter.api.Test;

/** Direct coverage for assertion-step validation and diagnostics. */
class AssertionStepTest {
  @Test
  void constructsAssertionStepsAndKeepsIncompatibleTargetsForStaticValidation() {
    Assertion assertion =
        new CellAssertion.Value(new dev.erst.gridgrind.core.model.value.CellValue.Text("Owner"));
    AssertionStep step =
        new AssertionStep(
            new StepId("assert-owner"),
            new CellSelector.ByAddress(new SheetName("Budget"), new CellAddress("A1")),
            assertion);

    assertEquals("assert-owner", step.stepId().value());
    assertEquals("ASSERTION", step.stepKind());
    assertEquals(assertion, step.assertion());

    assertDoesNotThrow(
        () -> new AssertionStep(new StepId("bad"), new WorkbookSelector.Current(), assertion));
  }
}
