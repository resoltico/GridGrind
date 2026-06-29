package dev.erst.gridgrind.contract.catalog;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.contract.step.WorkbookStep;
import org.junit.jupiter.api.Test;

/** Direct coverage for discriminator lookup failure wording. */
class GridGrindProtocolTypeNamesTest {
  @Test
  void rootTypesAreNotValidSubtypeLookups() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () -> GridGrindProtocolTypeNames.workbookStepTypeName(WorkbookStep.class));

    assertEquals(
        "No discriminator name registered for WorkbookStep subtype interface"
            + " dev.erst.gridgrind.contract.step.WorkbookStep",
        failure.getMessage());
  }
}
