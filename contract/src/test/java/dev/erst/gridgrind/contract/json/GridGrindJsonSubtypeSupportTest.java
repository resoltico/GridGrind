package dev.erst.gridgrind.contract.json;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

import com.fasterxml.jackson.annotation.JsonTypeName;
import org.junit.jupiter.api.Test;

/** Coverage for reflective sealed-subtype registration guards. */
class GridGrindJsonSubtypeSupportTest {
  @Test
  void namedLeafSubtypesRejectsLeafTypesWithoutJsonTypeNames() {
    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () -> GridGrindJsonSubtypeSupport.namedLeafSubtypes(MissingJsonTypeNameRoot.class));

    assertEquals(
        "Contract subtype "
            + MissingJsonTypeNameLeaf.class.getName()
            + " must declare a non-blank @JsonTypeName or @ProtocolTypeMetadata id",
        failure.getMessage());
  }

  @Test
  void namedLeafSubtypesRejectsBlankJsonTypeNames() {
    IllegalStateException failure =
        assertThrows(
            IllegalStateException.class,
            () -> GridGrindJsonSubtypeSupport.namedLeafSubtypes(BlankJsonTypeNameRoot.class));

    assertEquals(
        "Contract subtype "
            + BlankJsonTypeNameLeaf.class.getName()
            + " must declare a non-blank @JsonTypeName or @ProtocolTypeMetadata id",
        failure.getMessage());
  }

  /** Synthetic sealed root used to cover missing subtype annotations. */
  private sealed interface MissingJsonTypeNameRoot permits MissingJsonTypeNameLeaf {}

  /** Synthetic leaf without a JsonTypeName annotation. */
  private record MissingJsonTypeNameLeaf() implements MissingJsonTypeNameRoot {}

  /** Synthetic sealed root used to cover blank subtype annotations. */
  private sealed interface BlankJsonTypeNameRoot permits BlankJsonTypeNameLeaf {}

  /** Synthetic leaf with a blank JsonTypeName annotation. */
  @JsonTypeName("   ")
  private record BlankJsonTypeNameLeaf() implements BlankJsonTypeNameRoot {}
}
