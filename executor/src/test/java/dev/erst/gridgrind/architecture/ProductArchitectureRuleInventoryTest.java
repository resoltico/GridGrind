package dev.erst.gridgrind.architecture;

import static org.junit.jupiter.api.Assertions.assertEquals;

import org.junit.jupiter.api.Test;

/** Preserves the reviewed product architecture rule count. */
class ProductArchitectureRuleInventoryTest {
  @Test
  void retainsEveryMandatoryProductArchitectureRule() {
    assertEquals(16, ProductArchitectureRules.mandatoryRules().size());
  }
}
