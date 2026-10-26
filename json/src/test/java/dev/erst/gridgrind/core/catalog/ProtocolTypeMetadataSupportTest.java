package dev.erst.gridgrind.core.catalog;

import static org.junit.jupiter.api.Assertions.*;

import dev.erst.gridgrind.core.metadata.Union;
import dev.erst.gridgrind.core.metadata.Variant;
import java.util.List;
import org.junit.jupiter.api.Test;

/** Verifies the metadata-owned identity and sealed-alternative contract. */
class ProtocolTypeMetadataSupportTest {
  @Test
  void discoversDeclaredRecordAlternativesAndRejectsMissingIdentity() {
    assertEquals("FIRST", ProtocolTypeMetadataSupport.requiredTypeId(First.class));
    assertEquals(
        List.of(First.class, Second.class),
        ProtocolTypeMetadataSupport.subtypesForType(Choice.class));
    assertEquals(
        First.class,
        ProtocolTypeMetadataSupport.subtypeForTypeId(Choice.class, "FIRST").orElseThrow());
    assertTrue(ProtocolTypeMetadataSupport.subtypeForTypeId(Choice.class, "MISSING").isEmpty());
    assertThrows(
        IllegalStateException.class,
        () -> ProtocolTypeMetadataSupport.requiredTypeId(Unnamed.class));
    assertThrows(
        IllegalArgumentException.class,
        () -> ProtocolTypeMetadataSupport.subtypesForType(String.class));
  }

  /** Closed fixture whose identities are declared only on its records. */
  @Union
  private sealed interface Choice permits First, Second {}

  @Variant(id = "FIRST", summary = "First test choice.")
  private record First() implements Choice {}

  @Variant(id = "SECOND", summary = "Second test choice.")
  private record Second() implements Choice {}

  /** Deliberately missing the required variant identity. */
  private record Unnamed() {}
}
