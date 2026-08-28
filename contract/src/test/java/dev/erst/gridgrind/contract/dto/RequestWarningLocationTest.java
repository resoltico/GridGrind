package dev.erst.gridgrind.contract.dto;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import org.junit.jupiter.api.Test;

/** Direct coverage for request warning location ordering and byte-offset validation. */
class RequestWarningLocationTest {
  @Test
  void ordersOnlyStepWarningsAndRejectsNegativeByteOffsets() {
    assertEquals(7, new RequestWarningLocation.Step(7, "step", "SET_CELL").orderingStepIndex());
    assertEquals(-1, new RequestWarningLocation.RequestByteOffset(42).orderingStepIndex());
    assertThrows(
        IllegalArgumentException.class, () -> new RequestWarningLocation.RequestByteOffset(-1));
  }
}
