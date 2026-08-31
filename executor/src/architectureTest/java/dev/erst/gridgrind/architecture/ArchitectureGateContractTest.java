package dev.erst.gridgrind.architecture;

import static org.junit.jupiter.api.Assertions.assertEquals;

import com.tngtech.archunit.ArchConfiguration;
import org.junit.jupiter.api.Test;

/** Verifies the runtime facts that make the architecture gate fail closed. */
class ArchitectureGateContractTest {
  @Test
  void architectureTestRuntimeLoadsFailClosedPolicyAndEveryMandatoryRule() {
    assertEquals("true", ArchConfiguration.get().getProperty("archRule.failOnEmptyShould"));
    assertEquals(13, ProductArchitectureRules.mandatoryRules().size());
  }
}
