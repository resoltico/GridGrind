package dev.erst.gridgrind.core.model.validation;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import dev.erst.gridgrind.core.model.style.ColorSpec;
import dev.erst.gridgrind.core.vocabulary.RgbHex;
import org.junit.jupiter.api.Test;

/** Verifies the inclusive data-bar width bounds published by the protocol catalog. */
class ConditionalFormattingDataBarBoundsTest {
  @Test
  void rejectsWidthsAboveOneHundredAtEitherBound() {
    assertEquals(1.0d, new ConditionalFormattingThresholdSpec.Numeric(1.0d).value());
    assertThrows(IllegalArgumentException.class, () -> newDataBarRule(101, 101));
    assertThrows(IllegalArgumentException.class, () -> newDataBarRule(0, 101));
  }

  private static ConditionalFormattingRuleSpec.DataBarRule newDataBarRule(
      int widthMin, int widthMax) {
    return new ConditionalFormattingRuleSpec.DataBarRule(
        false,
        ColorSpec.rgb(new RgbHex("#112233")),
        false,
        widthMin,
        widthMax,
        new ConditionalFormattingThresholdSpec.Numeric(0.0d),
        new ConditionalFormattingThresholdSpec.Numeric(1.0d));
  }
}
