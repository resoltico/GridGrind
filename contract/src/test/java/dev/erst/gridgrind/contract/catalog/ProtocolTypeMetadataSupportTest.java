package dev.erst.gridgrind.contract.catalog;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

import com.fasterxml.jackson.annotation.JsonSubTypes;
import com.fasterxml.jackson.annotation.JsonTypeName;
import dev.erst.gridgrind.contract.selector.SheetSelector;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import org.junit.jupiter.api.Test;

/** Edge-path coverage for colocated protocol leaf metadata discovery. */
class ProtocolTypeMetadataSupportTest {
  @Test
  void requiredTypeIdSupportsJsonTypeNamesAndMetadataFallback() {
    assertEquals("JSON_TYPED", ProtocolTypeMetadataSupport.requiredTypeId(JsonTypeNamedLeaf.class));
    assertEquals(
        "METADATA_TYPED", ProtocolTypeMetadataSupport.requiredTypeId(MetadataOnlyLeaf.class));
    assertEquals(
        "METADATA_TYPED",
        ProtocolTypeMetadataSupport.requiredTypeId(BlankJsonTypeNameWithMetadataLeaf.class));

    IllegalStateException missingMetadataFailure =
        assertThrows(
            IllegalStateException.class,
            () -> ProtocolTypeMetadataSupport.requiredTypeId(MissingMetadataLeaf.class));
    assertEquals(
        "Contract subtype "
            + MissingMetadataLeaf.class.getName()
            + " must declare a non-blank @JsonTypeName or @ProtocolTypeMetadata id",
        missingMetadataFailure.getMessage());

    IllegalStateException blankMetadataIdFailure =
        assertThrows(
            IllegalStateException.class,
            () -> ProtocolTypeMetadataSupport.requiredTypeId(BlankMetadataIdLeaf.class));
    assertEquals(
        "Contract subtype "
            + BlankMetadataIdLeaf.class.getName()
            + " must declare a non-blank @JsonTypeName or @ProtocolTypeMetadata id",
        blankMetadataIdFailure.getMessage());
  }

  @Test
  void metadataLookupRejectsMissingBlankAndMismatchedMetadata() {
    IllegalStateException missingMetadataFailure =
        assertThrows(
            IllegalStateException.class,
            () -> ProtocolTypeMetadataSupport.requiredMetadata(MissingMetadataLeaf.class));
    assertEquals(
        "Contract subtype "
            + MissingMetadataLeaf.class.getName()
            + " must declare @ProtocolTypeMetadata",
        missingMetadataFailure.getMessage());

    IllegalStateException blankIdFailure =
        assertThrows(
            IllegalStateException.class,
            () -> ProtocolTypeMetadataSupport.requiredMetadata(BlankMetadataIdLeaf.class));
    assertEquals(
        "Contract subtype "
            + BlankMetadataIdLeaf.class.getName()
            + " must declare a non-blank metadata id",
        blankIdFailure.getMessage());

    IllegalStateException blankSummaryFailure =
        assertThrows(
            IllegalStateException.class,
            () -> ProtocolTypeMetadataSupport.requiredMetadata(BlankMetadataSummaryLeaf.class));
    assertEquals(
        "Contract subtype "
            + BlankMetadataSummaryLeaf.class.getName()
            + " must declare a non-blank metadata summary",
        blankSummaryFailure.getMessage());

    IllegalStateException mismatchFailure =
        assertThrows(
            IllegalStateException.class,
            () -> ProtocolTypeMetadataSupport.requiredMetadata(MismatchedIdsLeaf.class));
    assertEquals(
        "Contract subtype "
            + MismatchedIdsLeaf.class.getName()
            + " declares mismatched ids: wire=JSON_TYPED, metadata=METADATA_TYPED",
        mismatchFailure.getMessage());
  }

  @Test
  void selectorMetadataHelpersCoverStaticAndDynamicLeaves() {
    assertEquals(
        ProtocolTargetingMode.STATIC, ProtocolTypeMetadataSupport.targetingMode(StaticLeaf.class));
    assertEquals(
        Optional.empty(), ProtocolTypeMetadataSupport.targetSelectorRule(StaticLeaf.class));
    assertArrayEquals(
        new Class<?>[] {SheetSelector.class},
        ProtocolTypeMetadataSupport.staticTargetSelectors(StaticLeaf.class));

    assertEquals(
        ProtocolTargetingMode.ANALYSIS_QUERY,
        ProtocolTypeMetadataSupport.targetingMode(DynamicLeaf.class));
    assertEquals(
        Optional.of("dynamic selector rule"),
        ProtocolTypeMetadataSupport.targetSelectorRule(DynamicLeaf.class));

    IllegalArgumentException dynamicFailure =
        assertThrows(
            IllegalArgumentException.class,
            () -> ProtocolTypeMetadataSupport.staticTargetSelectors(DynamicLeaf.class));
    assertEquals(
        "Protocol subtype "
            + DynamicLeaf.class.getName()
            + " derives target selectors dynamically: dynamic selector rule",
        dynamicFailure.getMessage());
  }

  @Test
  void subtypeDiscoverySupportsJsonSubtypeRootsAndSealedRecordDescriptors() {
    assertEquals(
        Map.of(JsonSubtypeLeaf.class, "JSON_SUBTYPE"),
        ProtocolTypeMetadataSupport.typeIdsByClass(JsonSubtypeRoot.class));

    List<CatalogTypeDescriptor> descriptors =
        ProtocolTypeMetadataSupport.catalogDescriptorsFor(SealedCatalogRoot.class);
    assertEquals(1, descriptors.size());
    assertEquals("STATIC_TYPED", descriptors.getFirst().id());
    assertEquals("Static summary", descriptors.getFirst().summary());
    assertEquals(List.of("optionalField"), descriptors.getFirst().optionalFields());
  }

  @Test
  void subtypeDiscoveryRejectsRootsAndLeavesWithoutValidCatalogShapes() {
    IllegalArgumentException missingSubtypeDiscoveryFailure =
        assertThrows(
            IllegalArgumentException.class,
            () -> ProtocolTypeMetadataSupport.typeIdsByClass(PlainRoot.class));
    assertEquals(
        PlainRoot.class + " must be sealed or declare @JsonSubTypes for discriminator discovery",
        missingSubtypeDiscoveryFailure.getMessage());

    IllegalArgumentException unsealedCatalogFailure =
        assertThrows(
            IllegalArgumentException.class,
            () -> ProtocolTypeMetadataSupport.catalogDescriptorsFor(PlainRoot.class));
    assertEquals(PlainRoot.class + " must be sealed", unsealedCatalogFailure.getMessage());

    IllegalStateException nonRecordLeafFailure =
        assertThrows(
            IllegalStateException.class,
            () -> ProtocolTypeMetadataSupport.catalogDescriptorsFor(NonRecordCatalogRoot.class));
    assertTrue(
        nonRecordLeafFailure
            .getMessage()
            .contains("Catalog descriptor generation requires record leaf types"));
  }

  @JsonTypeName("JSON_TYPED")
  private record JsonTypeNamedLeaf() {}

  @ProtocolTypeMetadata(id = "METADATA_TYPED", summary = "Metadata summary")
  private record MetadataOnlyLeaf() {}

  @JsonTypeName("   ")
  @ProtocolTypeMetadata(id = "METADATA_TYPED", summary = "Metadata summary")
  private record BlankJsonTypeNameWithMetadataLeaf() {}

  private record MissingMetadataLeaf() {}

  @ProtocolTypeMetadata(id = " ", summary = "Metadata summary")
  private record BlankMetadataIdLeaf() {}

  @ProtocolTypeMetadata(id = "BLANK_SUMMARY", summary = " ")
  private record BlankMetadataSummaryLeaf() {}

  @JsonTypeName("JSON_TYPED")
  @ProtocolTypeMetadata(id = "METADATA_TYPED", summary = "Mismatch summary")
  private record MismatchedIdsLeaf() {}

  @ProtocolTypeMetadata(
      id = "STATIC_TYPED",
      summary = "Static summary",
      optionalFields = {"optionalField"},
      targetSelectors = {SheetSelector.class})
  private record StaticLeaf() {}

  @ProtocolTypeMetadata(
      id = "DYNAMIC_TYPED",
      summary = "Dynamic summary",
      targetingMode = ProtocolTargetingMode.ANALYSIS_QUERY,
      targetSelectorRule = "dynamic selector rule")
  private record DynamicLeaf() {}

  /** JsonSubTypes-backed discovery root used for non-sealed subtype lookup coverage. */
  @JsonSubTypes(@JsonSubTypes.Type(value = JsonSubtypeLeaf.class, name = "JSON_SUBTYPE"))
  private interface JsonSubtypeRoot {}

  /** JsonSubTypes-backed leaf used for non-sealed subtype lookup coverage. */
  private record JsonSubtypeLeaf() implements JsonSubtypeRoot {}

  @ProtocolTypeMetadata(
      id = "STATIC_TYPED",
      summary = "Static summary",
      optionalFields = {"optionalField"},
      targetSelectors = {SheetSelector.class})
  private record SealedCatalogLeaf(String optionalField) implements SealedCatalogRoot {}

  /** Sealed root used for positive catalog-descriptor discovery coverage. */
  private sealed interface SealedCatalogRoot permits SealedCatalogLeaf {}

  /** Sealed root used to reject catalog descriptors for non-record leaves. */
  private sealed interface NonRecordCatalogRoot permits NonRecordCatalogLeaf {}

  /** Non-record sealed leaf used to reject catalog descriptor generation. */
  @ProtocolTypeMetadata(
      id = "NON_RECORD_TYPED",
      summary = "Non-record summary",
      targetSelectors = {SheetSelector.class})
  private static final class NonRecordCatalogLeaf implements NonRecordCatalogRoot {}

  /** Plain non-sealed root used to reject descriptor discovery without subtype metadata. */
  private interface PlainRoot {}
}
