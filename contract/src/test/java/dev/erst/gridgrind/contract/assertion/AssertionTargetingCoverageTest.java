package dev.erst.gridgrind.contract.assertion;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import java.util.List;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Direct edge-path coverage for assertion-owned selector targeting helpers. */
class AssertionTargetingCoverageTest {
  @Test
  void rejectsUnmappedAssertionTypes() {
    @SuppressWarnings("unchecked")
    Class<? extends Assertion> unmappedAssertionType = (Class<? extends Assertion>) Assertion.class;

    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () -> Assertion.staticAllowedTargetTypesForType(unmappedAssertionType));

    assertEquals(
        "No target-type mapping configured for assertion class dev.erst.gridgrind.contract.assertion.Assertion",
        failure.getMessage());
    assertEquals(
        Optional.empty(), Assertion.dynamicTargetSelectorRuleForType(unmappedAssertionType));
    assertEquals(Optional.empty(), Assertion.targetSelectorRuleForType(unmappedAssertionType));
  }

  @Test
  void rejectsRecordTypesWithoutAssertionMetadataMappings() {
    @SuppressWarnings("unchecked")
    Class<? extends Assertion> missingMetadataAssertionType =
        (Class<? extends Assertion>) (Class<?>) MissingMetadataRecord.class;

    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () -> Assertion.staticAllowedTargetTypesForType(missingMetadataAssertionType));

    assertEquals(
        "Contract subtype "
            + MissingMetadataRecord.class.getName()
            + " must declare @ProtocolTypeMetadata",
        failure.getMessage());
  }

  @Test
  void rejectsEmptyCompositeAssertionFamilies() {
    IllegalArgumentException failure =
        assertThrows(
            IllegalArgumentException.class,
            () -> CompositeAssertion.commonTargetTypes(List.of(), "ANY_OF"));

    assertEquals(
        "ANY_OF requires nested assertions with compatible target families", failure.getMessage());
  }

  private record MissingMetadataRecord() {}
}
