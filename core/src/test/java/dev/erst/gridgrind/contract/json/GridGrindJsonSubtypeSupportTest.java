package dev.erst.gridgrind.contract.json;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeName;
import dev.erst.gridgrind.contract.catalog.ProtocolTypeMetadata;
import java.util.List;
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

  @Test
  void typeIdsReturnsNamesFromJsonSubTypesAnnotationFilteringBlanks() {
    List<String> ids = GridGrindJsonSubtypeSupport.typeIds(MixedJsonSubTypesRoot.class);

    assertEquals(List.of("NAMED_ONE"), ids);
  }

  @Test
  void typeIdsReturnsTypeNamesFromJsonTypeNameAnnotationOnSealedLeafs() {
    List<String> ids = GridGrindJsonSubtypeSupport.typeIds(JsonTypeNamedRoot.class);

    assertEquals(List.of("NAMED_LEAF"), ids);
  }

  @Test
  void typeIdsSkipsLeafsWithNoTypeAnnotations() {
    List<String> ids = GridGrindJsonSubtypeSupport.typeIds(MissingJsonTypeNameRoot.class);

    assertTrue(ids.isEmpty());
  }

  @Test
  void typeIdsSkipsLeafsWithBlankJsonTypeNameAndNoProtocolMetadata() {
    List<String> ids = GridGrindJsonSubtypeSupport.typeIds(BlankJsonTypeNameRoot.class);

    assertTrue(ids.isEmpty());
  }

  @Test
  void typeIdsSkipsLeafsWithBlankProtocolMetadataId() {
    List<String> ids = GridGrindJsonSubtypeSupport.typeIds(BlankProtocolMetadataRoot.class);

    assertTrue(ids.isEmpty());
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

  /** Synthetic root with @JsonSubTypes declaring one blank name and one valid name. */
  @JsonSubTypes({
    @JsonSubTypes.Type(value = UnnamedSubType.class, name = ""),
    @JsonSubTypes.Type(value = NamedSubType.class, name = "NAMED_ONE")
  })
  private interface MixedJsonSubTypesRoot {}

  private record UnnamedSubType() implements MixedJsonSubTypesRoot {}

  private record NamedSubType() implements MixedJsonSubTypesRoot {}

  /** Synthetic sealed root where the leaf uses @JsonTypeName. */
  private sealed interface JsonTypeNamedRoot permits JsonTypeNamedLeaf {}

  @JsonTypeName("NAMED_LEAF")
  private record JsonTypeNamedLeaf() implements JsonTypeNamedRoot {}

  /** Synthetic sealed root where the leaf has a @ProtocolTypeMetadata with a blank id. */
  private sealed interface BlankProtocolMetadataRoot permits BlankProtocolMetadataLeaf {}

  @ProtocolTypeMetadata(id = "   ", summary = "blank id leaf")
  private record BlankProtocolMetadataLeaf() implements BlankProtocolMetadataRoot {}
}
