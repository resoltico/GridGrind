package dev.erst.gridgrind.contract.assertion;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import java.util.List;
import org.junit.jupiter.api.Test;

/** Direct edge-path coverage for assertion-owned selector targeting helpers. */
class AssertionTargetingCoverageTest {
  @Test
  void rejectsEmptyCompositeAssertionFamilies() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () -> CompositeAssertion.commonTargetTypes(List.of(), "ANY_OF"));

    assertEquals(
        "ANY_OF requires nested assertions with compatible target families", failure.getMessage());
  }
}
