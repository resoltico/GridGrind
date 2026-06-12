package dev.erst.gridgrind.contract.assertion;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import org.junit.jupiter.api.Test;

/** Focused residual coverage for low-level assertion validation helpers. */
class AssertionSupportBehaviorTest {
  @Test
  void validatesNonBlankStringsAndFiniteNumbers() {
    assertEquals("Budget", AssertionSupport.requireNonBlank("Budget", "field"));
    assertEquals(12.5d, AssertionSupport.requireFiniteNumber(12.5d, "threshold"));

    assertEquals(
        "field must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> AssertionSupport.requireNonBlank(" ", "field"))
            .getMessage());
    assertEquals(
        "threshold must not be null",
        assertThrows(
                NullPointerException.class,
                () -> AssertionSupport.requireFiniteNumber(null, "threshold"))
            .getMessage());
    assertEquals(
        "threshold must be finite",
        assertThrows(
                IllegalArgumentException.class,
                () -> AssertionSupport.requireFiniteNumber(Double.NaN, "threshold"))
            .getMessage());
  }
}
