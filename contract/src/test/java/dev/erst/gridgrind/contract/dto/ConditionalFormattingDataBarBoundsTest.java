package dev.erst.gridgrind.contract.dto;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import org.junit.jupiter.api.Test;

/** Verifies the inclusive data-bar width bounds published by the protocol catalog. */
class ConditionalFormattingDataBarBoundsTest {
  @Test
  void rejectsWidthsAboveOneHundredAtEitherBound() {
    assertEquals(1.0d, new ConditionalFormattingThresholdInput.Numeric(1.0d).value());
    assertThrows(IllegalArgumentException.class, () -> newDataBarRule(101, 101));
    assertThrows(IllegalArgumentException.class, () -> newDataBarRule(0, 101));
  }

  private static ConditionalFormattingRuleInput.DataBarRule newDataBarRule(
      int widthMin, int widthMax) {
    return new ConditionalFormattingRuleInput.DataBarRule(
        false,
        ColorInput.rgb("#112233"),
        false,
        widthMin,
        widthMax,
        new ConditionalFormattingThresholdInput.Numeric(0.0d),
        new ConditionalFormattingThresholdInput.Numeric(1.0d));
  }
}
