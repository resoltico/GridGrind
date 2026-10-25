package dev.erst.gridgrind.contract.dto;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;

import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Verifies RGB normalization rejects blank input and preserves canonical uppercase output. */
class ProtocolRgbColorSupportTest {
  @Test
  void normalizesRequiredAndOptionalRgbValues() {
    String lowerCaseRgb = "#a1b2c3";
    String upperCaseRgb = "#A1B2C3";

    assertNotNull(ProtocolRgbColorSupport.newForVerification());
    assertEquals(
        Optional.of(upperCaseRgb), ProtocolRgbColorSupport.normalizeRgbHex(lowerCaseRgb, "color"));
    assertEquals(upperCaseRgb, ProtocolRgbColorSupport.requireRgbHex(lowerCaseRgb, "color"));
    assertEquals(
        "color must not be blank",
        assertThrows(
                IllegalArgumentException.class,
                () -> ProtocolRgbColorSupport.normalizeRgbHex(" ", "color"))
            .getMessage());
    assertThrows(
        IllegalArgumentException.class,
        () -> ProtocolRgbColorSupport.normalizeRgbHex("#a1b2cg", "color"));
  }
}
