package dev.erst.gridgrind.core.vocabulary;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import org.junit.jupiter.api.Test;

/** Verifies RGB normalization rejects blank input and preserves canonical uppercase output. */
class RgbHexTest {
  @Test
  void normalizesAndValidatesRgbLiterals() {
    String lowerCaseRgb = "#a1b2c3";
    String upperCaseRgb = "#A1B2C3";

    assertEquals(upperCaseRgb, new RgbHex(lowerCaseRgb).value());
    assertEquals(
        "rgb must not be blank",
        assertThrows(IllegalArgumentException.class, () -> new RgbHex(" ")).getMessage());
    assertThrows(IllegalArgumentException.class, () -> new RgbHex("#a1b2cg"));
  }
}
