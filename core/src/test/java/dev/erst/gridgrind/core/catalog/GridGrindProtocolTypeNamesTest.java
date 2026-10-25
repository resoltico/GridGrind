package dev.erst.gridgrind.core.catalog;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.plan.WorkbookStep;
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
            + " dev.erst.gridgrind.core.plan.WorkbookStep",
        failure.getMessage());
  }
}
