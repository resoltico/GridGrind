package dev.erst.gridgrind.core.plan;

import static org.junit.jupiter.api.Assertions.*;

import org.junit.jupiter.api.Test;

/** Verifies real selector identities without a parallel subtype list. */
class SelectorMetadataTest {
  @Test
  void selectorIdentityComesFromItsDeclaredVariants() {
    assertEquals("CELL_BY_ADDRESS", SelectorTypes.typeIdFor(CellSelector.ByAddress.class));
    assertEquals(
        CellSelector.ByAddress.class, SelectorTypes.typeFor("CELL_BY_ADDRESS").orElseThrow());
    assertTrue(SelectorTypes.typeFor("UNKNOWN").isEmpty());
    assertTrue(SelectorTypes.typeIdsFor(CellSelector.class).contains("CELL_BY_ADDRESS"));
    assertThrows(IllegalArgumentException.class, () -> SelectorTypes.typeIdFor(String.class));
  }
}
